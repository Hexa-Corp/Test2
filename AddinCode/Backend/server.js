import express from 'express';
import axios from 'axios';
import cors from 'cors';

const app = express();

// Allow requests from localhost:3000 (your add-in)
app.use(cors({
  origin: 'https://localhost:3000'
}));

const tenantId = "2ef4acb6-d902-4515-8021-6eeeeb5d12bc";
const clientId = "e48c8505-c759-40a7-909e-7d2228c584ad";
const clientSecret = "V5y8Q~ALhSPnh64IirN6taHQE5pxzk2SU6T6vai6";

app.get('/getAppToken', async (req, res) => {
  try {
    const response = await axios.post(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials"
      })
    );
    res.json({ token: response.data.access_token });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(3001, () => console.log("Server running on port 3001"));




/* global Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Office is ready!");
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    document.getElementById("run").addEventListener("click", showSharePointDataRaw);
  }
});

async function showSharePointDataRaw() {
  try {
    // 1️⃣ Get app-only token from backend
    const tokenRes = await fetch("http://localhost:3001/getAppToken");
    const tokenData = await tokenRes.json();
    const token = tokenData.token;

    // 2️⃣ Fetch SharePoint list items
    const listUrl = "https://graph.microsoft.com/v1.0/sites/tylky.sharepoint.com,76aa5505-18b0-420c-aac0-fb41b3f15a31,304365ac-39cb-4017-9892-b2642829d467/lists/0f565f3b-e4be-4dd9-ad16-9803cbee809f/items?expand=fields";
    const res = await fetch(listUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json"
      }
    });

    if (!res.ok) throw new Error(`Graph API error: ${res.status} ${res.statusText}`);

    const data = await res.json();

    // 3️⃣ Insert raw JSON into Word
    await Word.run(async (context) => {
      const docBody = context.document.body;

      // Clear previous content
      docBody.clear();

      // Insert JSON string (pretty-printed)
      const jsonString = JSON.stringify(data.value, null, 2);
      docBody.insertParagraph(jsonString, Word.InsertLocation.start);

      await context.sync();
    });

  } catch (err) {
    console.error("Error fetching SharePoint list:", err);
    alert("Failed to load SharePoint list. See console for details.");
  }
}
