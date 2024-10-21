require("dotenv").config();
const express = require("express");
const msal = require("@azure/msal-node");
const axios = require("axios");
const cors = require("cors");
const { cwd } = require("process");
const fs = require("fs");
const path = require("path");
const app = express();
app.use(express.json());
app.use(cors());

const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/common`,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
  },
};
const cca = new msal.ConfidentialClientApplication(msalConfig);

app.get("/auth/login", async (req, res) => {
  const authUrl = cca.getAuthCodeUrl({
    scopes: ["user.read", "files.read"],
    redirectUri: process.env.REDIRECT_URI,
    prompt: "login",
  });
  authUrl
    .then((url) => {
      console.log(url);
      res.redirect(url);
    })
    .catch((err) => res.status(500).json({ error: err.message }));
});

app.post("/auth/callback", (req, res) => {
  const tokenRequest = {
    code: req.body.code,
    scopes: ["user.read", "files.read"],
    redirectUri: process.env.REDIRECT_URI,
  };

  cca
    .acquireTokenByCode(tokenRequest)
    .then((response) => {
      const accessToken = response.accessToken;
      const expiresOn = 600; //10mins
      res.status(200).json({ accessToken, expiresOn });
    })
    .catch((err) => res.status(500).json({ error: err.message }));
});

app.get("/me", (req, res) => {
  const accessToken = req.query.token;
  axios
    .get("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${accessToken}` },
    })
    .then((response) => {
      res.json(response.data);
    })
    .catch((err) => res.status(500).json({ error: err.message }));
});

app.get("/onedrive/files", (req, res) => {
  const accessToken = req.query.token;
  axios
    .get("https://graph.microsoft.com/v1.0/me/drive/root/children", {
      headers: { Authorization: `Bearer ${accessToken}` },
    })
    .then((response) => {
      res.json(response.data);
    })
    .catch((err) => res.status(500).json({ error: err.message }));
});

app.get("/onedrive/download/:fileId", async (req, res) => {
  const token = req.query.token;
  const { fileId } = req.params;

  try {
    const fileResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
        },
      }
    );

    const downloadUrl = fileResponse.data["@microsoft.graph.downloadUrl"];
    const fileName = fileResponse.data.name;
    const fileStream = await axios({
      url: downloadUrl,
      method: "GET",
      responseType: "stream",
    });

    const filePath = path.join(cwd(), "public", fileName);
    const publicDir = path.join(cwd(), "public");
    if (!fs.existsSync(publicDir)) {
      fs.mkdirSync(publicDir);
    }
    const writeStream = fs.createWriteStream(filePath);
    fileStream.data.pipe(writeStream);

    writeStream.on("finish", () => {
      return res.status(200).json({
        message: "File downloaded and saved",
        filePath: `/public/${fileName}`,
      });
    });

    writeStream.on("error", (error) => {
      console.error("Error saving file:", error);
      return res.status(500).json({ error: "Failed to save file" });
    });
  } catch (error) {
    console.error("Error downloading the file:", error);
    return res.status(500).json({ error: "Failed to download file" });
  }
});

app.post("/download-files", async (req, res) => {
  const token = req.query.token;
  const { fileIds } = req.body;

  if (!fileIds || fileIds.length === 0) {
    return res.status(400).json({ error: "No files selected." });
  }

  try {
    const publicDir = path.join(cwd(), "public");
    if (!fs.existsSync(publicDir)) {
      fs.mkdirSync(publicDir);
    }

    for (const fileId of fileIds) {
      const fileResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`,
        {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      );

      const downloadUrl = fileResponse.data["@microsoft.graph.downloadUrl"];
      const fileName = fileResponse.data.name;
      const response = await axios.get(downloadUrl, {
        responseType: "stream",
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });

      const filePath = path.join(publicDir, fileName);
      const writer = fs.createWriteStream(filePath);
      response.data.pipe(writer);
      await new Promise((resolve, reject) => {
        writer.on("finish", resolve);
        writer.on("error", reject);
      });
    }
    res.status(200).json({ message: "Files downloaded successfully." });
  } catch (error) {
    console.error("Error downloading files:", error);
    res.status(500).json({ error: "Failed to download files." });
  }
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
