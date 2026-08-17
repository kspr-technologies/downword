Great question! Migrating from **Postgres 12** to **Postgres 16** is mostly
mechanical, but there are three or four places where people get burned. Here's
how I'd approach it.

## TL;DR

- Use `pg_upgrade --link` if you can afford the downtime — it's ~10x faster.
- **Do not** skip the `ANALYZE` afterwards; the planner will be blind until you do.
- Watch out for the `standard_conforming_strings` change if you're coming from
  a *very* old dump.

---

## 1. Pre-flight checks

First, confirm the extensions you actually use:

```sql
SELECT extname, extversion FROM pg_extension ORDER BY extname;
```

Then check for anything deprecated:

| Extension | PG 12 | PG 16 | Action |
|-----------|:-----:|:-----:|--------|
| `postgis` | 3.0 | 3.4 | Upgrade **before** the server upgrade |
| `pg_stat_statements` | ✅ | ✅ | No action |
| `chkpass` | ✅ | ❌ | **Removed** — migrate to `pgcrypto` |
| `adminpack` | ✅ | ❌ | Drop it; nothing uses it |

> ⚠️ **Important:** if `chkpass` is in use, the upgrade will *abort* with
> `could not load library`. Check first.

## 2. The upgrade itself

1. Stop the old cluster:

   ```bash
   sudo systemctl stop postgresql@12-main
   ```

2. Run the compatibility check (this does **not** modify anything):

   ```bash
   /usr/lib/postgresql/16/bin/pg_upgrade \
     --old-datadir=/var/lib/postgresql/12/main \
     --new-datadir=/var/lib/postgresql/16/main \
     --old-bindir=/usr/lib/postgresql/12/bin \
     --new-bindir=/usr/lib/postgresql/16/bin \
     --check
   ```

3. If that passes, drop the `--check` and run it for real. With `--link`:

   ```bash
   ... --link
   ```

   > **Note:** `--link` hard-links the data files, so **the old cluster becomes
   > unusable**. Take a backup you can actually restore from first.

4. Rebuild statistics — this is the step everyone forgets:

   ```bash
   /usr/lib/postgresql/16/bin/vacuumdb --all --analyze-in-stages
   ```

## 3. Things that changed and will bite you

- [x] `pg_dump` now emits `SET transaction_timeout` — harmless, but it breaks
      naive `grep`-based dump diffing.
- [x] The default `password_encryption` is `scram-sha-256`. Old clients
      (`libpq` < 10, JDBC < 42.2) **cannot connect**.
- [ ] `LOGICAL` replication slots do *not* survive `pg_upgrade`. Recreate them.
- [ ] `SEARCH_PATH` handling in `SECURITY DEFINER` functions is stricter.

### Client compatibility

If you see this:

```
FATAL:  password authentication failed for user "app"
DETAIL:  User "app" does not have a valid SCRAM secret.
```

...it means the stored hash is still MD5. Fix it with:

```sql
ALTER SYSTEM SET password_encryption = 'scram-sha-256';
SELECT pg_reload_conf();
-- then, per user:
\password app
```

## 4. Rollback plan

| Scenario | Recovery | RTO |
| -------- | -------- | --: |
| `--check` fails | Nothing to undo | 0 |
| Upgrade fails mid-way, no `--link` | Start the old cluster | ~1 min |
| Upgrade fails mid-way, with `--link` | Restore from backup | ~45 min |

---

**Let me know** which Postgres distribution you're on (Debian packages? RDS?
Docker?) and I can tailor the commands — RDS in particular does *none* of this
the same way.
