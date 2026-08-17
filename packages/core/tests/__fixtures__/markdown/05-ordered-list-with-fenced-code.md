Here's how to set it up:

1. Install the dependencies:

   ```bash
   npm install express cors
   ```

2. Create a `server.js` file:

   ```js
   const express = require("express");
   const app = express();

   app.get("/", (req, res) => {
     res.json({ ok: true });
   });

   app.listen(3000);
   ```

3. Run it:

   ```
   node server.js
   ```

   You should see the server start on port 3000.

4. Verify with `curl`:

   ```bash
   curl -s localhost:3000 | jq .
   ```

   > **Note:** if `jq` isn't installed, drop the pipe.

That's it! Let me know if you hit any errors.
