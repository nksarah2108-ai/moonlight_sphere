const express = require("express");
const cors = require("cors");
const { spawn } = require("child_process");
const fs = require("fs");
const path = require("path");

const app = express();

app.use(cors());
app.use(express.json());
app.use(express.static("public")); // untuk index.html

/* =====================================
   TEST ROUTE
===================================== */
app.get("/", (req, res) => {
  res.send("Server RPH berjalan dengan baik ✅");
});

/* =====================================
   BM - GENERATE RPH
===================================== */
app.post("/generate-rph", (req, res) => {

  const { minggu, tarikh, kelas, hari_dipilih } = req.body;

  console.log("=== DATA BM DITERIMA ===");
  console.log(req.body);

  const python = spawn("python3", [
    "generate.py",
    minggu,
    tarikh,
    kelas,
    JSON.stringify(hari_dipilih)
  ]);

  python.stdout.on("data", (data) => {
    console.log("PYTHON BM:", data.toString());
  });

  python.stderr.on("data", (data) => {
    console.log("PYTHON BM ERROR:", data.toString());
  });

  python.on("close", (code) => {
    console.log("BM exit code:", code);

    if (code === 0) {
      res.json({ success: true });
    } else {
      res.json({ error: "Gagal jana RPH BM." });
    }
  });

});


/* =====================================
   RBT - GENERATE
===================================== */
app.post("/generate-rbt", (req, res) => {

  const { minggu, tarikh, hari, kelas, masa, refleksi } = req.body;

  console.log("=== DATA RBT DITERIMA ===");
  console.log(req.body);

  const python = spawn("python", [
    "generate_rbt_t5.py",
    minggu,
    tarikh,
    hari,
    kelas,
    masa,
    refleksi
  ]);

  python.stdout.on("data", (data) => {
    console.log("PYTHON RBT:", data.toString());
  });

  python.stderr.on("data", (data) => {
    console.log("PYTHON RBT ERROR:", data.toString());
  });

  python.on("close", (code) => {
    console.log("RBT exit code:", code);

    if (code === 0) {
      res.json({ success: true });
    } else {
      res.json({ error: "Gagal jana RBT." });
    }
  });

});

/* =====================================
   DOWNLOAD RBT
===================================== */
app.get("/download-rbt", (req, res) => {

  const folderPath = path.join(__dirname, "output");

  if (!fs.existsSync(folderPath)) {
    return res.send("Folder output tidak wujud.");
  }

  const files = fs.readdirSync(folderPath)
    .filter(file => file.endsWith(".docx"));

  if (files.length === 0) {
    return res.send("Tiada fail RBT untuk dimuat turun.");
  }

  // Ambil file paling latest
  const latestFile = files.sort().reverse()[0];

  const filePath = path.join(folderPath, latestFile);

  res.download(filePath);
});
/* =====================================
   DOWNLOAD BM
===================================== */
app.get("/download-rph", (req, res) => {

  const filePath = path.join(__dirname, "RPH_BM_FINAL_OUTPUT.pptx");

  if (!fs.existsSync(filePath)) {
    return res.send("Fail BM belum dijana.");
  }

  res.download(filePath);
});

/* =====================================
   START SERVER
===================================== */
const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
