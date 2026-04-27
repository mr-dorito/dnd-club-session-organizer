import express from "express";
import cors from "cors";
import OpenAI from "openai";

const app = express();
const port = Number(process.env.PORT || 3001);
const allowedOrigin = process.env.ALLOWED_ORIGIN || "*";
const model = process.env.OPENAI_MODEL || "gpt-5.2";

app.use(cors({ origin: allowedOrigin }));
app.use(express.json({ limit: "1mb" }));

const recapSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    summary: { type: "string" },
    events: { type: "string" },
    npcs: { type: "string" },
    locations: { type: "string" },
    quests: { type: "string" },
    unresolvedThreads: { type: "string" },
  },
  required: ["summary", "events", "npcs", "locations", "quests", "unresolvedThreads"],
};

function getOpenAIClient() {
  if (!process.env.OPENAI_API_KEY) {
    return null;
  }
  return new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

function normalizeRecapPayload(payload) {
  return {
    summary: String(payload?.summary || "").trim(),
    events: String(payload?.events || "").trim(),
    npcs: String(payload?.npcs || "").trim(),
    locations: String(payload?.locations || "").trim(),
    quests: String(payload?.quests || "").trim(),
    unresolvedThreads: String(payload?.unresolvedThreads || "").trim(),
  };
}

app.get("/api/health", (_request, response) => {
  response.json({
    ok: true,
    service: "dnd-club-ai-backend",
    hasOpenAIKey: Boolean(process.env.OPENAI_API_KEY),
  });
});

app.post("/api/generate-recap", async (request, response) => {
  const transcript = String(request.body?.transcript || "").trim();
  if (!transcript) {
    response.status(400).json({ error: "Transcript is required before generating a recap." });
    return;
  }

  const client = getOpenAIClient();
  if (!client) {
    response.status(500).json({ error: "OPENAI_API_KEY is not configured on the backend." });
    return;
  }

  const campaignName = String(request.body?.campaignName || "Unknown campaign").trim();
  const sessionLabel = String(request.body?.sessionLabel || "Current session").trim();
  const existingRecap = normalizeRecapPayload(request.body?.existingRecap || {});

  try {
    const aiResponse = await client.responses.create({
      model,
      input: [
        {
          role: "system",
          content:
            "You help a Dungeon Master turn D&D session transcripts into concise, useful campaign notes. Return only the requested structured recap fields.",
        },
        {
          role: "user",
          content: [
            `Campaign: ${campaignName}`,
            `Session: ${sessionLabel}`,
            "",
            "Existing notes, if any:",
            JSON.stringify(existingRecap),
            "",
            "Transcript:",
            transcript,
          ].join("\n"),
        },
      ],
      text: {
        format: {
          type: "json_schema",
          name: "session_recap",
          strict: true,
          schema: recapSchema,
        },
      },
    });

    const rawText = aiResponse.output_text || "{}";
    const recap = normalizeRecapPayload(JSON.parse(rawText));
    response.json(recap);
  } catch (error) {
    console.error("AI recap generation failed:", error);
    response.status(500).json({
      error: "AI recap generation failed. Try again later or keep editing the notes manually.",
    });
  }
});

app.listen(port, () => {
  console.log(`D&D Club AI backend listening on port ${port}`);
});
