> Before you run any of this, note the following:
>
> 1. Take a backup:
>
>    ```bash
>    pg_dump -Fc mydb > mydb.dump
>    ```
>
> 2. Stop the writers.
>
>    - the API workers
>    - the cron jobs
>
> 3. Only then run the migration.
>
> > If step 1 fails, **stop**. Do not continue.
>
> | Phase | Duration |
> | ----- | -------- |
> | dump  | ~2 min   |
> | apply | ~30 s    |
