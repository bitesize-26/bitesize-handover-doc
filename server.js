import express from "express";
import cors from "cors";
import crypto from "crypto";
import fs from "fs";
import path from "path";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  ImageRun,
} from "docx";

const app = express();
app.use(cors());
app.use(express.json({ limit: "5mb" }));

// ===== Branding =====
const BRAND_PINK = "FF0087";
const FONT_PRIMARY = "Abel"; // Will only render as Abel if installed on the viewer's machine.
const LOGO_PATH = path.join(process.cwd(), "assets", "bitesize-logo.png");
const logoImage = fs.readFileSync(LOGO_PATH);

// ===== Temp download store (in-memory) =====
const TEMP_TTL_MS = 15 * 60 * 1000; // 15 minutes
const tempFiles = new Map();

function putTempFile({ buffer, mime, filename }) {
  const token = crypto.randomBytes(18).toString("hex");
  const expiresAt = Date.now() + TEMP_TTL_MS;
  tempFiles.set(token, { buffer, mime, filename, expiresAt });
  return token;
}

setInterval(() => {
  const now = Date.now();
  for (const [token, meta] of tempFiles.entries()) {
    if (!meta?.expiresAt || meta.expiresAt <= now) tempFiles.delete(token);
  }
}, 60 * 1000).unref();

// ===== Helpers =====
function labelRun(text) {
  return new TextRun({
    text,
    bold: true,
    color: "000000",
    font: FONT_PRIMARY,
  });
}

function valueRun(text) {
  return new TextRun({
    text: text ?? "",
    color: "000000",
    font: FONT_PRIMARY,
  });
}

function field(label, value) {
  return new Paragraph({
    children: [labelRun(`${label}: `), valueRun(value ?? "")],
    spacing: { after: 120 },
  });
}

function sectionTitle(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        bold: true,
        color: BRAND_PINK,
        font: FONT_PRIMARY,
        size: 28, // 14pt
      }),
    ],
    spacing: { before: 240, after: 120 },
  });
}

function divider() {
  return new Paragraph({
    children: [
      new TextRun({
        text: "______________________________________________________________",
        color: "000000",
        font: FONT_PRIMARY,
      }),
    ],
    spacing: { before: 80, after: 160 },
  });
}

function safeAsciiFilename(name) {
  // Keep downloads reliable everywhere (no fancy punctuation)
  const cleaned = String(name || "BiteSize")
    .replace(/[^\w\s.-]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
  return cleaned ? `${cleaned} - BiteSize Handover.docx` : "BiteSize-Handover.docx";
}

// ===== Routes =====
app.post("/generate-handover", async (req, res) => {
  try {
    const d = req.body;

    console.log("generate-handover hit:", {
      business_name: d?.business_name,
      client_name: d?.client_name,
    });

    if (!d?.business_name || !d?.client_name) {
      return res.status(400).json({ ok: false, error: "Missing required fields" });
    }

    const doc = new Document({
      sections: [
        {
          children: [
            // Logo
            new Paragraph({
              children: [
                new ImageRun({
                  data: logoImage,
                  transformation: { width: 220, height: 70 },
                }),
              ],
              spacing: { after: 160 },
            }),

            // Title
            new Paragraph({
              children: [
                new TextRun({
                  text: "SUB-BRAND MARKETING HANDOVER",
                  bold: true,
                  color: "000000",
                  font: FONT_PRIMARY,
                  size: 32, // 16pt
                }),
              ],
              spacing: { after: 240 },
            }),

            divider(),

            sectionTitle("Client Details"),
            field("Client", d.client_name),
            field("Business", d.business_name),
            field("Phone Number", d.phone_number),
            field("Email Address", d.email_address),

            divider(),

            sectionTitle("Commercial Agreement"),
            field("Tier", d.tier),
            field("Monthly Tier Fee", d.monthly_tier_fee),
            field("Contract Length", d.contract_length),
            field("Preferred Start Date", d.preferred_start_date),

            divider(),

            sectionTitle("Time & Revenue Summary"),
            field("Base Monthly Hours", d.base_monthly_hours),
            field("Add-On Recurring Hours", d.addon_recurring_hours),
            field("Total Recurring Monthly Hours", d.total_recurring_hours),
            field("Total Recurring Revenue", d.total_recurring_revenue),
            field("Effective Hourly Rate", d.effective_hourly_rate),

            divider(),

            sectionTitle("Project Add-Ons (Non-Recurring)"),
            field("Project Add-Ons", d.project_addons),
            field("Project Hours", d.project_hours),
            field("Project Revenue", d.project_revenue),

            divider(),

            sectionTitle("Delivery Scope"),
            field("Channels", d.channels),
            field("Posting Frequency", d.posting_frequency),
            field("Primary Objective", d.primary_objective),

            divider(),

            sectionTitle("Additional Notes"),
            field("Add-Ons Selected", d.addons_selected),
            field("Scope Exceptions", d.scope_exceptions),
            field("Assets Ready", d.assets_ready),

            divider(),

            sectionTitle("Risk & Governance"),
            field("Red Flags", d.red_flags),
            field("Risk Severity", d.risk_severity),
            field("Dropbox Folder Link", d.dropbox_link),
            field("Folder Name", d.folder_name),
            field("Assets Uploaded", d.assets_uploaded),
            field("Capacity Impact", d.capacity_impact),
          ],
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);

    const mime =
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    // Provide a friendly filename, but keep it ASCII-safe
    const filename = safeAsciiFilename(d.business_name);

    const token = putTempFile({ buffer, mime, filename });

    const baseUrl = `${req.protocol}://${req.get("host")}`;
    const download_url = `${baseUrl}/download/${token}`;

    return res.status(200).json({
      ok: true,
      filename,
      mime_type: mime,
      download_url,
      expires_in_seconds: Math.floor(TEMP_TTL_MS / 1000),
    });
  } catch (err) {
    console.error("generate-handover error:", err);
    return res.status(500).json({ ok: false, error: String(err?.message ?? err) });
  }
});

app.get("/download/:token", (req, res) => {
  const { token } = req.params;
  const meta = tempFiles.get(token);

  if (!meta) return res.status(404).send("File expired or not found.");

  if (meta.expiresAt <= Date.now()) {
    tempFiles.delete(token);
    return res.status(410).send("File expired.");
  }

  res.status(200);
  res.setHeader("Content-Type", meta.mime);
  res.setHeader("Content-Length", String(meta.buffer.length));

  // Keep header safe: fixed ASCII filename
  res.setHeader("Content-Disposition", `attachment; filename="${meta.filename}"`);

  return res.end(meta.buffer);
});

app.get("/health", (req, res) => res.json({ ok: true }));

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Server running on port ${port}`));
