import express from "express";
import cors from "cors";
import OpenAI, { toFile } from "openai";
import multer from "multer";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const app = express();
const port = Number(process.env.PORT || 3001);
const allowedOrigin = process.env.ALLOWED_ORIGIN || "*";
const model = process.env.OPENAI_MODEL || "gpt-5.2";
const transcriptionModel = process.env.TRANSCRIPTION_MODEL || "gpt-4o-transcribe";
const speakerTimingModel = process.env.SPEAKER_TIMING_MODEL || "gpt-4o-transcribe-diarize";
const maxAudioBytes = 25 * 1024 * 1024;
const allowedAudioExtensions = new Set([".mp3", ".mp4", ".mpeg", ".mpga", ".m4a", ".wav", ".webm"]);
const upload = multer({
  dest: path.join(os.tmpdir(), "dnd-club-audio"),
  limits: { fileSize: maxAudioBytes, files: 1 },
});

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

function getAudioExtension(fileName) {
  return path.extname(String(fileName || "")).toLowerCase();
}

async function deleteUploadedFile(file) {
  if (!file?.path) return;
  try {
    await fs.unlink(file.path);
  } catch (error) {
    if (error.code !== "ENOENT") {
      console.warn("Could not delete temporary audio upload:", error);
    }
  }
}

async function validateAudioUpload(file, response, emptyMessage) {
  if (!file) {
    response.status(400).json({ error: emptyMessage });
    return false;
  }

  const extension = getAudioExtension(file.originalname);
  if (!allowedAudioExtensions.has(extension)) {
    await deleteUploadedFile(file);
    response.status(400).json({
      error: "Unsupported audio type. Use mp3, mp4, mpeg, mpga, m4a, wav, or webm.",
    });
    return false;
  }

  return true;
}

function normalizeSpeakerSegments(payload) {
  const segments = Array.isArray(payload?.segments) ? payload.segments : [];
  return segments
    .map((segment, index) => {
      const startSeconds = Math.max(0, Number(segment.start ?? segment.startSeconds ?? segment.start_seconds ?? 0) || 0);
      const endSeconds = Math.max(
        startSeconds,
        Number(segment.end ?? segment.endSeconds ?? segment.end_seconds ?? startSeconds) || startSeconds,
      );
      const durationSeconds = Math.max(
        0,
        Number(segment.duration ?? segment.durationSeconds ?? segment.duration_seconds ?? endSeconds - startSeconds) || 0,
      );
      const speakerLabel = String(
        segment.speaker ?? segment.speakerLabel ?? segment.speaker_label ?? `Speaker ${index + 1}`,
      ).trim();
      return {
        id: `speaker-segment-${index + 1}`,
        speakerLabel: speakerLabel || `Speaker ${index + 1}`,
        text: String(segment.text || "").trim(),
        startSeconds,
        endSeconds,
        durationSeconds,
      };
    })
    .filter((segment) => segment.text || segment.durationSeconds > 0);
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

app.post("/api/transcribe-session-audio", upload.single("audio"), async (request, response) => {
  const file = request.file;
  if (!(await validateAudioUpload(file, response, "Choose an audio file before starting transcription."))) {
    return;
  }

  const client = getOpenAIClient();
  if (!client) {
    await deleteUploadedFile(file);
    response.status(500).json({ error: "OPENAI_API_KEY is not configured on the backend." });
    return;
  }

  try {
    const audioFile = await toFile(await fs.readFile(file.path), file.originalname, { type: file.mimetype });
    const transcriptResponse = await client.audio.transcriptions.create({
      file: audioFile,
      model: transcriptionModel,
      response_format: "json",
    });

    response.json({
      transcript: String(transcriptResponse.text || "").trim(),
      durationSeconds: Number(transcriptResponse.duration || 0) || null,
      fileName: file.originalname,
    });
  } catch (error) {
    console.error("Audio transcription failed:", error);
    response.status(500).json({
      error: "Audio transcription failed. Try a smaller file or keep the transcript manual for now.",
    });
  } finally {
    await deleteUploadedFile(file);
  }
});

app.post("/api/process-speaker-timing", upload.single("audio"), async (request, response) => {
  const file = request.file;
  if (!(await validateAudioUpload(file, response, "Choose an audio file before processing speaker timing."))) {
    return;
  }

  const client = getOpenAIClient();
  if (!client) {
    await deleteUploadedFile(file);
    response.status(500).json({ error: "OPENAI_API_KEY is not configured on the backend." });
    return;
  }

  try {
    const audioFile = await toFile(await fs.readFile(file.path), file.originalname, { type: file.mimetype });
    const transcriptResponse = await client.audio.transcriptions.create({
      file: audioFile,
      model: speakerTimingModel,
      response_format: "diarized_json",
    });
    const speakerSegments = normalizeSpeakerSegments(transcriptResponse);

    response.json({
      transcript: String(transcriptResponse.text || "").trim(),
      speakerSegments,
      durationSeconds: Number(transcriptResponse.duration || 0) || null,
      fileName: file.originalname,
    });
  } catch (error) {
    console.error("Speaker timing failed:", error);
    response.status(500).json({
      error: "Speaker timing failed. Try normal transcription or keep speaking time manual for now.",
    });
  } finally {
    await deleteUploadedFile(file);
  }
});

app.use((error, _request, response, next) => {
  if (!error) {
    next();
    return;
  }
  if (error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE") {
    response.status(413).json({ error: "Audio files must be 25 MB or smaller for this transcription step." });
    return;
  }
  response.status(500).json({ error: "The backend could not process that request." });
});

app.listen(port, () => {
  console.log(`D&D Club AI backend listening on port ${port}`);
});
