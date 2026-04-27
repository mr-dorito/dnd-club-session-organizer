# D&D Club AI Backend Deployment

## Render Setup

1. Push this project to GitHub.
2. In Render, create a new Blueprint or Web Service from the repo.
3. Use `render.yaml`, or set these values manually:
   - Root directory: `server`
   - Build command: `npm install`
   - Start command: `npm start`
4. Add environment variables:
   - `OPENAI_API_KEY`: your OpenAI API key
   - `OPENAI_MODEL`: `gpt-5.2`
   - `ALLOWED_ORIGIN`: `*`

## Connect The App

1. Copy the Render service URL after deployment.
2. Open the app and go to `Session Notes`.
3. Paste the URL into `Backend URL`.
4. Click `Save backend URL`.
5. Add a transcript and click `Generate recap from transcript`.

## Local Testing

From the `server` folder:

```bash
npm install
OPENAI_API_KEY=your_openai_api_key_here npm start
```

The local backend runs at:

```text
http://localhost:3001
```

Health check:

```text
http://localhost:3001/api/health
```

The app automatically uses `http://localhost:3001` when opened from localhost. When opened as a `file://` page, paste `http://localhost:3001` into the app's `Backend URL` field.
