import express from "express";
import cors from "cors";
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from "docx";

const app = express();
app.use(cors());
app.use(express.json({ limit: "2mb" }));

function field(label, value) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${label}: `, bold: true }),
      new TextRun({ text: value ?? "" }),
    ],
    spacing: { after: 120 },
  });
}

function sectionTitle(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
  });
}

function divider() {
  return new Paragraph({
    children: [new TextRun({ text: "______________________________________________________________" })],
    spacing: { before: 80, after: 160 },
  });
}

app.post("/generate-handover", async (req, res) => {
  try {
    const d = req.body;

    if (!d?.business_name || !d?.client_name) {
      return res.status(400).json({ error: "Missing required fields" });
    }

    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              text: "SUB-BRAND MARKETING HANDOVER",
              heading: HeadingLevel.HEADING_1,
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

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${d.business_name} – BiteSize Handover.docx"`
    );

    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

app.get("/health", (req, res) => res.json({ ok: true }));

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Server running on port ${port}`));
