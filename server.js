import express from "express";
import multer from "multer";
import axios from "axios";
import FormData from "form-data";
import cors from "cors";
import path from "path";

const app = express();
const upload = multer();

// 生产环境CORS：允许Vercel域名
const allowedOrigins = [
  "https://wiseaudit.vercel.app",           // 你的Vercel域名
  "https://wiseaudit-*.vercel.app",         // Vercel预览域名
  "http://localhost:3000",                   // 本地开发
  "http://localhost:5173",
];

app.use(cors({
  origin: function (origin, callback) {
    // 允许无origin的请求（如Postman）
    if (!origin) return callback(null, true);
    if (allowedOrigins.some(o => {
      if (o.includes('*')) {
        const regex = new RegExp(o.replace('*', '.*'));
        return regex.test(origin);
      }
      return o === origin;
    })) {
      return callback(null, true);
    }
    callback(new Error("Not allowed by CORS"));
  },
  credentials: true,
  methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization", "X-Requested-With"]
}));

// 解析JSON请求体
app.use(express.json());

// Python后端地址（Render内部网络或环境变量）
const PYTHON_API_BASE = process.env.PYTHON_API_URL || "http://127.0.0.1:8000";
const AUDIT_API = `${PYTHON_API_BASE}/audit/ar/run`;

console.log(`[Backend] Python API地址: ${PYTHON_API_BASE}`);
console.log(`[Backend] 环境: ${process.env.NODE_ENV || 'development'}`);

/**
 * =========================
 * 1️⃣ 审计执行接口
 * =========================
 */
app.post("/api/run-audit", upload.any(), async (req, res) => {
  try {
    const form = new FormData();

    console.log("========== Node收到文件 ==========");
    req.files.forEach((file) => {
      console.log(file.fieldname, "=>", file.originalname);
      form.append(file.fieldname, file.buffer, file.originalname);
    });

    const response = await axios.post(AUDIT_API, form, {
      headers: form.getHeaders(),
      maxBodyLength: Infinity,
      timeout: 120000,
    });

    const result = response.data;

    console.log("========== Python返回 ==========");
    console.log(result);

    /**
     * 生产环境：改写下载地址为当前域名
     */
    if (result.download_url) {
      const filename = result.download_url.split("/").pop();
      // 使用当前服务的域名
      const host = req.headers.host || process.env.RENDER_EXTERNAL_HOSTNAME || 'localhost';
      const protocol = req.headers['x-forwarded-proto'] || 'https';
      result.download_full_url = `${protocol}://${host}/api/download/${filename}`;
    }

    res.json(result);

  } catch (err) {
    console.error("❌ Node后端错误：", err.response?.data || err.message);

    res.status(500).json({
      message: err.message,
      detail: err.response?.data || null,
    });
  }
});


/**
 * =========================
 * 2️⃣ 下载底稿代理
 * =========================
 */
app.get("/api/download/:filename", async (req, res) => {
  try {
    const { filename } = req.params;

    console.log("📥 请求下载文件：", filename);

    const fileUrl = `${PYTHON_API_BASE}/audit/ar/download/${filename}`;

    const response = await axios.get(fileUrl, {
      responseType: "stream",
    });

    // 设置下载头
    res.setHeader(
      "Content-Type",
      response.headers["content-type"] ||
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      response.headers["content-disposition"] ||
      `attachment; filename="${filename}"`
    );

    // 流式返回文件
    response.data.pipe(res);

  } catch (err) {
    console.error("❌ 下载代理失败：", err.response?.data || err.message);

    res.status(500).json({
      message: "下载底稿失败",
      detail: err.response?.data || null,
    });
  }
});


/**
 * =========================
 * 3️⃣ 健康检查接口
 * =========================
 */
app.get("/api/health", async (req, res) => {
  // 同时检查Python后端状态
  let pythonStatus = 'unknown';
  try {
    await axios.get(`${PYTHON_API_BASE}/`, { timeout: 3000 });
    pythonStatus = 'online';
  } catch {
    pythonStatus = 'offline';
  }

  res.json({
    status: "ok",
    message: "Node backend is running",
    pythonBackend: pythonStatus,
    timestamp: new Date().toISOString(),
    version: "1.0.0-cloud"
  });
});


/**
 * =========================
 * 4️⃣ 启动服务
 * =========================
 */
const PORT = process.env.PORT || 3001;
app.listen(PORT, "0.0.0.0", () => {
  console.log(`🚀 Node后端已启动：0.0.0.0:${PORT}`);
  console.log(`📡 健康检查：http://0.0.0.0:${PORT}/api/health`);
});
