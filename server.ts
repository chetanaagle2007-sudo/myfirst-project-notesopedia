import express from "express";
import path from "path";
import dotenv from "dotenv";
import fs from "fs";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI } from "@google/genai";

// Load environment variables
dotenv.config();

const app = express();
const PORT = 3000;

app.use(express.json());

// Persistent user notes file path
const NOTES_FILE_PATH = path.join(process.cwd(), "user_notes.json");

// Default template notes to populate the repository initially
const DEFAULT_NOTES = [
  {
    id: "note-1",
    title: "Deep Dive into ACID Transactions & Concurrency Control",
    content: "ACID properties ensure relational databases handle errors gracefully. \n\n### Atomicity\nAll changes in a transaction must complete successfully, or all are rolled back. This is managed via the **Undo Log**.\n\n### Consistency\nEnforces rules (constraints, keys, check constraints) from one valid state to another.\n\n### Isolation\nEnsures concurrent execution of transactions leaves the database in the same state as if they were executed sequentially. Isolated levels include Read Uncommitted, Read Committed, Repeatable Read, and Serializable.\n\n### Durability\nGuarantees that committed transaction data is permanently written to disk, typically achieved using **Write-Ahead Logging (WAL)** or transaction logs so it survives power losses.",
    subjectName: "Relational Database Management Systems",
    subjectCode: "2505MJCT201",
    topicName: "ACID Properties & Recovery",
    uploaderName: "Prof. Rajesh Kumar",
    uploaderRole: "Professor",
    uploaderEmail: "rkumar@university.edu",
    uploadedAt: "2026-06-29T10:30:00.000Z",
    likes: 12
  },
  {
    id: "note-2",
    title: "SQL Indexing Performance: B-Tree vs Hash Indexes",
    content: "When should you use which index? Let's analyze the disk IO performance:\n\n1. **B-Tree Indexes** (Default in PostgreSQL/Oracle):\n   - **Time Complexity**: O(log N) for Search, Insert, Delete.\n   - **Best Used For**: Range queries (e.g., `WHERE age BETWEEN 20 AND 30`), equality lookup, and sorted results (`ORDER BY`).\n   - **Structure**: Balanced tree with high fanout, keeping depth low (usually 3-4 levels) to minimize disk page fetches.\n\n2. **Hash Indexes**:\n   - **Time Complexity**: O(1) average lookup.\n   - **Best Used For**: Strictly equality comparison queries (e.g., `WHERE status = 'ACTIVE'`).\n   - **Limitations**: Cannot be used for range queries or index-sorted retrieval.\n\n*Pro-Tip*: PostgreSQL B-Trees also support index-only scans when selecting columns fully covered by the index structure.",
    subjectName: "Database Systems & Administration",
    subjectCode: "CS-302",
    topicName: "Index Optimization & Disk Storage",
    uploaderName: "Chetana Agle",
    uploaderRole: "Lead TA",
    uploaderEmail: "chetanaagle2007@gmail.com",
    uploadedAt: "2026-06-30T04:15:00.000Z",
    likes: 24
  },
  {
    id: "note-3",
    title: "Normal Forms Quick Sheet (1NF to BCNF)",
    content: "Use this simple checklist for normalization exam questions:\n\n- **1NF**: Atomic values only. No repeating columns or array-like values.\n- **2NF**: In 1NF + **No partial dependency** (non-key columns must depend on the *entire* primary key, not a subset of a composite key).\n- **3NF**: In 2NF + **No transitive dependency** (non-key columns cannot depend on other non-key columns).\n- **BCNF**: Must be in 3NF + For every functional dependency `A -> B`, `A` must be a super key.",
    subjectName: "Relational Database Management Systems",
    subjectCode: "2505MJCT201",
    topicName: "Database Normalization",
    uploaderName: "Amit Sharma",
    uploaderRole: "Student",
    uploaderEmail: "amit.sharma99@student.edu",
    uploadedAt: "2026-06-30T07:45:00.000Z",
    likes: 7
  }
];

// Helper to read notes
function readNotesFromFile(): any[] {
  try {
    if (fs.existsSync(NOTES_FILE_PATH)) {
      const data = fs.readFileSync(NOTES_FILE_PATH, "utf8");
      return JSON.parse(data);
    } else {
      // Create file with defaults
      fs.writeFileSync(NOTES_FILE_PATH, JSON.stringify(DEFAULT_NOTES, null, 2), "utf8");
      return DEFAULT_NOTES;
    }
  } catch (err) {
    console.error("Error reading notes file:", err);
    return DEFAULT_NOTES;
  }
}

// Helper to write notes
function writeNotesToFile(notes: any[]) {
  try {
    fs.writeFileSync(NOTES_FILE_PATH, JSON.stringify(notes, null, 2), "utf8");
  } catch (err) {
    console.error("Error writing notes file:", err);
  }
}

// API Routes for User Notes
app.get("/api/notes", (req, res) => {
  const notes = readNotesFromFile();
  res.json(notes);
});

app.post("/api/notes", (req, res) => {
  try {
    const { title, content, subjectName, subjectCode, topicName, uploaderName, uploaderRole, uploaderEmail } = req.body;
    
    if (!title || !content || !subjectName || !uploaderName) {
      return res.status(400).json({ error: "Missing required note details (title, content, subject, uploader name)" });
    }

    const notes = readNotesFromFile();
    const newNote = {
      id: "note-" + Date.now(),
      title,
      content,
      subjectName,
      subjectCode: subjectCode || "GEN-DB",
      topicName: topicName || "General Database Concept",
      uploaderName,
      uploaderRole: uploaderRole || "Student",
      uploaderEmail: uploaderEmail || "",
      uploadedAt: new Date().toISOString(),
      likes: 0
    };

    notes.unshift(newNote); // Put newest notes at the top
    writeNotesToFile(notes);

    res.status(201).json(newNote);
  } catch (err: any) {
    console.error("Error adding note:", err);
    res.status(500).json({ error: "Failed to upload your study note." });
  }
});

app.post("/api/notes/:id/like", (req, res) => {
  try {
    const { id } = req.params;
    const notes = readNotesFromFile();
    const noteIndex = notes.findIndex(n => n.id === id);
    if (noteIndex !== -1) {
      notes[noteIndex].likes = (notes[noteIndex].likes || 0) + 1;
      writeNotesToFile(notes);
      return res.json(notes[noteIndex]);
    }
    res.status(404).json({ error: "Note not found" });
  } catch (err) {
    res.status(500).json({ error: "Failed to like note" });
  }
});

// GLOBAL SYSTEM CONFIGURATION & CONTROLS FOR THE APP
const CONFIG_FILE_PATH = path.join(process.cwd(), "system_config.json");
const DEFAULT_CONFIG = {
  announcement: "🎓 Welcome to Notsopedia Hub! Check out the real-time PL/SQL simulator and peer study archives.",
  announcementActive: true,
  enableSimulator: true,
  enableSubmissions: true
};

function readConfigFromFile(): any {
  try {
    if (fs.existsSync(CONFIG_FILE_PATH)) {
      const data = fs.readFileSync(CONFIG_FILE_PATH, "utf8");
      return JSON.parse(data);
    } else {
      fs.writeFileSync(CONFIG_FILE_PATH, JSON.stringify(DEFAULT_CONFIG, null, 2), "utf8");
      return DEFAULT_CONFIG;
    }
  } catch (err) {
    return DEFAULT_CONFIG;
  }
}

function writeConfigToFile(config: any) {
  try {
    fs.writeFileSync(CONFIG_FILE_PATH, JSON.stringify(config, null, 2), "utf8");
  } catch (err) {
    console.error("Error writing config file:", err);
  }
}

// Get global config settings
app.get("/api/system/config", (req, res) => {
  res.json(readConfigFromFile());
});

// Update global config settings (Allows controlling the app)
app.post("/api/system/config", (req, res) => {
  try {
    const current = readConfigFromFile();
    const updated = { ...current, ...req.body };
    writeConfigToFile(updated);
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: "Failed to update configuration settings" });
  }
});

// Delete note Admin endpoint
app.delete("/api/notes/:id", (req, res) => {
  try {
    const { id } = req.params;
    const notes = readNotesFromFile();
    const filtered = notes.filter(n => n.id !== id);
    if (notes.length !== filtered.length) {
      writeNotesToFile(filtered);
      return res.json({ success: true, id });
    }
    res.status(404).json({ error: "Note not found" });
  } catch (err) {
    res.status(500).json({ error: "Failed to delete note" });
  }
});

// Edit/Update note Admin endpoint
app.put("/api/notes/:id", (req, res) => {
  try {
    const { id } = req.params;
    const { title, content, subjectName, subjectCode, topicName, uploaderName, uploaderRole, uploaderEmail } = req.body;
    
    const notes = readNotesFromFile();
    const noteIndex = notes.findIndex(n => n.id === id);
    if (noteIndex !== -1) {
      notes[noteIndex] = {
        ...notes[noteIndex],
        title: title || notes[noteIndex].title,
        content: content || notes[noteIndex].content,
        subjectName: subjectName || notes[noteIndex].subjectName,
        subjectCode: subjectCode || notes[noteIndex].subjectCode,
        topicName: topicName || notes[noteIndex].topicName,
        uploaderName: uploaderName || notes[noteIndex].uploaderName,
        uploaderRole: uploaderRole || notes[noteIndex].uploaderRole,
        uploaderEmail: uploaderEmail !== undefined ? uploaderEmail : notes[noteIndex].uploaderEmail
      };
      writeNotesToFile(notes);
      return res.json(notes[noteIndex]);
    }
    res.status(404).json({ error: "Note not found" });
  } catch (err) {
    res.status(500).json({ error: "Failed to update note" });
  }
});

// Reset notes to default state
app.post("/api/notes/reset", (req, res) => {
  try {
    writeNotesToFile(DEFAULT_NOTES);
    res.json(DEFAULT_NOTES);
  } catch (err) {
    res.status(500).json({ error: "Failed to reset database" });
  }
});

// Lazy-initialized Gemini Client
let aiClient: GoogleGenAI | null = null;

function getGemini(): GoogleGenAI {
  if (!aiClient) {
    const key = process.env.GEMINI_API_KEY;
    if (!key) {
      throw new Error("GEMINI_API_KEY environment variable is required. Please set it in Settings > Secrets.");
    }
    aiClient = new GoogleGenAI({
      apiKey: key,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        }
      }
    });
  }
  return aiClient;
}

// API Health Check
app.get("/api/health", (req, res) => {
  res.json({ status: "ok", time: new Date().toISOString() });
});

// AI Tutor Endpoint for Explanations & Custom Answers
app.post("/api/tutor/explain", async (req, res) => {
  try {
    const { topic, format, additionalPrompt } = req.body;
    if (!topic) {
      return res.status(400).json({ error: "Topic is required" });
    }

    const ai = getGemini();
    
    let prompt = "";
    if (format === "exam-answer") {
      prompt = `Act as an expert RDBMS professor and textbook author. The user is preparing for their university theory exam tomorrow.
Provide a complete, full-marks, 5-mark exam-style answer for the topic: "${topic}".
You MUST write the response in this exact order and format (using clear Markdown formatting, headings, bullet points, and code/syntax boxes where applicable):

1. **Definition**: A clear, concise, 2-3 sentence definition.
2. **Syntax (if applicable)**: Provide the correct SQL/PLSQL syntax inside a formatted code block. If not applicable to this topic, provide an architectural or mathematical model description.
3. **Types / Features**: A structured list of sub-types or main features, explained in detail.
4. **Advantages / Importance**: A clean bulleted list of advantages of using this concept.
5. **Practical Example**: Provide a complete, fully-functional, and realistic SQL or PL/pgSQL example matching standard PostgreSQL or RDBMS environments. Keep it correct, illustrative, and clean.

Additional constraints or requests: ${additionalPrompt || "None"}.
Ensure the content is deep, professional, and directly ready to be written in a university exam.`;
    } else if (format === "quiz-viva") {
      prompt = `Act as an RDBMS Viva-Voce Examiner. Generate a set of 5 highly expected viva/interview questions for the topic or unit: "${topic}".
For each question, provide:
1. The question.
2. The expected "ideal" answer that will impress the external examiner.
3. A quick "pro tip" or keyword they must mention.
Format nicely in Markdown with clear headings or bulleted structures.`;
    } else {
      prompt = `Provide a detailed revision note for: "${topic}". Keep it highly structured, professional, and clear. Incorporate diagrams or flowcharts using ascii/text drawings if helpful. Additional instruction: ${additionalPrompt || "none"}`;
    }

    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
    });

    res.json({ text: response.text });
  } catch (error: any) {
    console.error("Gemini API Error:", error);
    res.status(500).json({ error: error.message || "Failed to generate AI response" });
  }
});

// AI Evaluate Answer Endpoint (for mock tests)
app.post("/api/tutor/evaluate", async (req, res) => {
  try {
    const { question, userAnswer } = req.body;
    if (!question || !userAnswer) {
      return res.status(400).json({ error: "Question and User Answer are required" });
    }

    const ai = getGemini();

    const prompt = `Act as an RDBMS university paper evaluator. A student has submitted an answer to the following question:
Question: "${question}"
Student's Answer: "${userAnswer}"

Analyze their answer and provide a highly constructive, detailed feedback report. Format the output in JSON format with the following keys:
- "score": A score out of 5 (integer, e.g. 1 to 5). Be realistic but encouraging.
- "verdict": A 1-sentence verdict (e.g., "Excellent start, but missing syntax details").
- "strengths": A string array listing what they did well.
- "gaps": A string array listing what they missed (e.g., failed to explain ACID fully, missing semicolon in trigger syntax).
- "modelAnswer": The perfect full-marks exam-style answer containing: Definition, Syntax, Types, Advantages, and an Example.

Return ONLY a valid JSON object matching this schema. Do not enclose it in markdown code blocks like \`\`\`json. Just raw JSON string.`;

    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
      }
    });

    const resultText = response.text ? response.text.trim() : "{}";
    res.json(JSON.parse(resultText));
  } catch (error: any) {
    console.error("Gemini API Error:", error);
    res.status(500).json({ error: error.message || "Failed to evaluate answer" });
  }
});

// Setup Vite or Production Static Serving
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
