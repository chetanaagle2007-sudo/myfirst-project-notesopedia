import React, { useState, useEffect } from "react";
import pptxgen from "pptxgenjs";
import { 
  BookOpen, 
  Terminal, 
  Award, 
  Brain, 
  Play, 
  RefreshCw, 
  UserCheck, 
  ChevronRight, 
  HelpCircle,
  FileText,
  Search,
  CheckCircle,
  AlertCircle,
  Sparkles,
  Lock,
  ArrowRight,
  Maximize2,
  UploadCloud,
  Plus,
  Heart,
  User,
  Folder,
  Tag,
  Mail,
  Shield,
  Trash2,
  Edit3,
  Check
} from "lucide-react";
import { STUDY_UNITS, EXPECTED_QUESTIONS, VIVA_QUESTIONS, MEMORY_TRICKS } from "./data";
import { StudyTopic, StudyUnit, ExpectedQuestion, VivaQuestion, UserNote } from "./types";

export default function App() {
  // General Navigation State
  const [activeTab, setActiveTab] = useState<"syllabus" | "simulator" | "mock-exam" | "viva" | "notesopedia">("syllabus");
  const [selectedUnitId, setSelectedUnitId] = useState<string>("unit-2-plsql");
  const [selectedTopicId, setSelectedTopicId] = useState<string>("plsql-overview");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [isGeneratingPPTX, setIsGeneratingPPTX] = useState<boolean>(false);

  // Notsopedia states
  const [userNotes, setUserNotes] = useState<UserNote[]>([]);
  const [isLoadingNotes, setIsLoadingNotes] = useState<boolean>(false);
  const [notesSearchQuery, setNotesSearchQuery] = useState<string>("");
  const [selectedSubjectFilter, setSelectedSubjectFilter] = useState<string>("All Subjects");
  const [expandedNoteId, setExpandedNoteId] = useState<string | null>(null);

  // Admin & System Control States
  const [isAdminMode, setIsAdminMode] = useState<boolean>(false);
  const [adminPassword, setAdminPassword] = useState<string>("");
  const [showAdminPasswordModal, setShowAdminPasswordModal] = useState<boolean>(false);
  const [editingNote, setEditingNote] = useState<UserNote | null>(null);
  const [isSavingConfig, setIsSavingConfig] = useState<boolean>(false);
  const [systemConfig, setSystemConfig] = useState<{
    announcement: string;
    announcementActive: boolean;
    enableSimulator: boolean;
    enableSubmissions: boolean;
  }>({
    announcement: "🎓 Welcome to Notsopedia Hub! Check out the real-time PL/SQL simulator and peer study archives.",
    announcementActive: true,
    enableSimulator: true,
    enableSubmissions: true
  });

  // Notes Form State
  const [newNoteTitle, setNewNoteTitle] = useState<string>("");
  const [newNoteContent, setNewNoteContent] = useState<string>("");
  const [newNoteSubjectName, setNewNoteSubjectName] = useState<string>("Relational Database Management Systems");
  const [newNoteSubjectCode, setNewNoteSubjectCode] = useState<string>("2505MJCT201");
  const [newNoteTopicName, setNewNoteTopicName] = useState<string>("");
  const [newNoteUploaderName, setNewNoteUploaderName] = useState<string>("");
  const [newNoteUploaderRole, setNewNoteUploaderRole] = useState<string>("Student");
  const [newNoteUploaderEmail, setNewNoteUploaderEmail] = useState<string>("");
  
  const [isSubmittingNote, setIsSubmittingNote] = useState<boolean>(false);
  const [showUploadForm, setShowUploadForm] = useState<boolean>(false);
  const [notesError, setNotesError] = useState<string>("");
  const [notesSuccessMessage, setNotesSuccessMessage] = useState<string>("");
  const [pptxStatus, setPptxStatus] = useState<string>("");
  const [pptxError, setPptxError] = useState<string>("");
  const [pptxBase64, setPptxBase64] = useState<string>("");

  // Dynamic PowerPoint Presentation Builder
  const handleDownloadPPTX = async () => {
    setIsGeneratingPPTX(true);
    setPptxStatus("Initializing PowerPoint library...");
    setPptxError("");
    setPptxBase64("");
    try {
      // Robust constructor resolution for dual ESM/CJS environments
      let pptxConstructor: any = pptxgen;
      if (!pptxConstructor) {
        throw new Error("PowerPoint library failed to load. Check installation.");
      }
      if (typeof pptxConstructor !== "function" && (pptxConstructor as any).default) {
        pptxConstructor = (pptxConstructor as any).default;
      }
      if (typeof pptxConstructor !== "function") {
        throw new Error("PowerPoint library export structure is unexpected.");
      }

      setPptxStatus("Creating new presentation layout...");
      const pres = new pptxConstructor();
      pres.layout = "LAYOUT_169";

      // 1. Cover Slide (High Contrast Theme)
      const slide1 = pres.addSlide();
      slide1.background = { color: "E4E3E0" };
      
      // Black top Accent Box
      slide1.addShape("rect", {
        x: 0,
        y: 0,
        w: 13.33,
        h: 0.4,
        fill: { color: "141414" }
      });

      slide1.addText("RDBMS EXAM MASTERCLASS", {
        x: 1.0,
        y: 1.8,
        w: 11.3,
        h: 1.2,
        fontSize: 34,
        fontFace: "Georgia",
        italic: true,
        bold: true,
        color: "141414",
      });

      slide1.addText("Comprehensive Lecture Slides & Expected University Exam Solutions", {
        x: 1.0,
        y: 3.1,
        w: 11.3,
        h: 0.5,
        fontSize: 16,
        fontFace: "Arial",
        color: "141414",
      });

      slide1.addText("Course Code: 2505MJCT201  •  Time: 2 Hours  •  Max Marks: 30", {
        x: 1.0,
        y: 4.5,
        w: 11.3,
        h: 0.4,
        fontSize: 12,
        fontFace: "Courier New",
        bold: true,
        color: "141414",
      });

      slide1.addText("Generated via AI Study Companion", {
        x: 1.0,
        y: 6.2,
        w: 5.0,
        h: 0.4,
        fontSize: 10,
        fontFace: "Arial",
        color: "666666",
      });

      // Black bottom Accent Box
      slide1.addShape("rect", {
        x: 0,
        y: 7.1,
        w: 13.33,
        h: 0.4,
        fill: { color: "141414" }
      });

      // 2. Add Slides for each Study Unit and its child Topics
      setPptxStatus("Building slide contents for syllabus units...");
      STUDY_UNITS.forEach((unit) => {
        // Section Break Slide
        const divider = pres.addSlide();
        divider.background = { color: "141414" };

        divider.addText(unit.title.toUpperCase(), {
          x: 1.2,
          y: 2.5,
          w: 11.0,
          h: 1.0,
          fontSize: 28,
          fontFace: "Georgia",
          italic: true,
          bold: true,
          color: "E4E3E0",
        });

        divider.addText(unit.description, {
          x: 1.2,
          y: 3.8,
          w: 11.0,
          h: 1.2,
          fontSize: 15,
          fontFace: "Arial",
          color: "CCCCCC",
        });

        // Add topic detail slides
        unit.topics.forEach((topic) => {
          const slide = pres.addSlide();
          slide.background = { color: "E4E3E0" };

          // Top Sub-header
          slide.addText(`${unit.title.split(":")[0].toUpperCase()} • TOPIC BLUEPRINT`, {
            x: 0.8,
            y: 0.4,
            w: 11.7,
            h: 0.3,
            fontSize: 10,
            fontFace: "Courier New",
            color: "666666",
            bold: true,
          });

          slide.addText(topic.title, {
            x: 0.8,
            y: 0.7,
            w: 11.7,
            h: 0.6,
            fontSize: 22,
            fontFace: "Georgia",
            bold: true,
            color: "141414",
          });

          // Horizontal division line
          slide.addShape("line", {
            x: 0.8,
            y: 1.3,
            w: 11.7,
            h: 0,
            line: { color: "141414", width: 1.5 }
          });

          // Left Box: Core Definition
          let leftContent = `CORE CONCEPT / DEFINITION:\n${topic.definition}\n\n`;
          if (topic.typesOrFeatures && topic.typesOrFeatures.length > 0) {
            leftContent += `KEY CLASSIFICATIONS & ATTRIBUTES:\n`;
            topic.typesOrFeatures.forEach((feat) => {
              leftContent += `• ${feat}\n`;
            });
          }
          slide.addText(leftContent, {
            x: 0.8,
            y: 1.6,
            w: 5.6,
            h: 4.8,
            fontSize: 11,
            fontFace: "Arial",
            color: "141414",
            valign: "top",
          });

          // Right Box: SQL Code Block or Advantages Box
          if (topic.example) {
            // High-contrast tech code container
            slide.addShape("rect", {
              x: 6.8,
              y: 1.6,
              w: 5.7,
              h: 4.8,
              fill: { color: "141414" },
            });

            slide.addText("EXAM ANSWER PATTERN: SQL SYNTAX", {
              x: 7.1,
              y: 1.8,
              w: 5.1,
              h: 0.3,
              fontSize: 10,
              fontFace: "Courier New",
              color: "00FF00",
              bold: true,
            });

            slide.addText(topic.example, {
              x: 7.1,
              y: 2.1,
              w: 5.1,
              h: 4.1,
              fontSize: 8.5,
              fontFace: "Courier New",
              color: "E4E3E0",
              valign: "top",
            });
          } else if (topic.advantages && topic.advantages.length > 0) {
            // Styled Advantages box
            slide.addShape("rect", {
              x: 6.8,
              y: 1.6,
              w: 5.7,
              h: 4.8,
              fill: { color: "FFFFFF" },
              line: { color: "141414", width: 1.5 },
            });

            slide.addText("EXAM WRITING PATTERN: ADVANTAGES", {
              x: 7.1,
              y: 1.8,
              w: 5.1,
              h: 0.3,
              fontSize: 11,
              fontFace: "Courier New",
              color: "141414",
              bold: true,
            });

            let advStr = "";
            topic.advantages.forEach((adv) => {
              advStr += `✔ ${adv}\n\n`;
            });

            slide.addText(advStr, {
              x: 7.1,
              y: 2.2,
              w: 5.1,
              h: 4.0,
              fontSize: 10.5,
              fontFace: "Arial",
              color: "141414",
              valign: "top",
            });
          }
        });
      });

      // 3. Expected 5-Mark Questions Slide
      setPptxStatus("Creating exam blueprint slides...");
      const eqSlide = pres.addSlide();
      eqSlide.background = { color: "141414" };

      eqSlide.addText("EXPECTED 5-MARK EXAM THEORY BLUEPRINT", {
        x: 0.8,
        y: 0.5,
        w: 11.7,
        h: 0.6,
        fontSize: 22,
        fontFace: "Georgia",
        bold: true,
        italic: true,
        color: "E4E3E0",
      });

      EXPECTED_QUESTIONS.forEach((eq, idx) => {
        const itemY = 1.3 + (idx * 1.05);
        eqSlide.addText(`Q${idx + 1}. ${eq.question} (5 Marks)`, {
          x: 0.8,
          y: itemY,
          w: 11.7,
          h: 0.4,
          fontSize: 12.5,
          fontFace: "Arial",
          bold: true,
          color: "FFFFFF",
        });

        eqSlide.addText(`EXAM WRITING STRATEGY: ${eq.hint}`, {
          x: 0.8,
          y: itemY + 0.35,
          w: 11.7,
          h: 0.4,
          fontSize: 10,
          fontFace: "Courier New",
          color: "00FF00",
        });
      });

      // 4. Memory Mnemonics Slide
      setPptxStatus("Generating acronyms and mnemonics...");
      const trickSlide = pres.addSlide();
      trickSlide.background = { color: "E4E3E0" };

      trickSlide.addText("LAST-MINUTE MEMORY ACRONYMS & MNEMONICS", {
        x: 0.8,
        y: 0.5,
        w: 11.7,
        h: 0.6,
        fontSize: 22,
        fontFace: "Georgia",
        bold: true,
        color: "141414",
      });

      trickSlide.addShape("line", {
        x: 0.8,
        y: 1.2,
        w: 11.7,
        h: 0,
        line: { color: "141414", width: 1.5 }
      });

      MEMORY_TRICKS.forEach((trick, idx) => {
        const colIdx = idx % 2;
        const rowIdx = Math.floor(idx / 2);
        const trickX = 0.8 + (colIdx * 6.0);
        const trickY = 1.5 + (rowIdx * 1.8);

        // Solid clean container box
        trickSlide.addShape("rect", {
          x: trickX,
          y: trickY,
          w: 5.5,
          h: 1.5,
          fill: { color: "FFFFFF" },
          line: { color: "141414", width: 1.5 },
        });

        trickSlide.addText(`${trick.acronym} (${trick.concept})`, {
          x: trickX + 0.2,
          y: trickY + 0.1,
          w: 5.1,
          h: 0.3,
          fontSize: 12,
          fontFace: "Courier New",
          bold: true,
          color: "141414",
        });

        trickSlide.addText(`Expansion: ${trick.expansion}\n\nMnemonic Key: ${trick.explanation}`, {
          x: trickX + 0.2,
          y: trickY + 0.4,
          w: 5.1,
          h: 1.0,
          fontSize: 9.5,
          fontFace: "Arial",
          color: "333333",
          valign: "top",
        });
      });

      // 5. Viva Tips & Interview Readiness
      setPptxStatus("Finalizing viva and Q&A slides...");
      const finalSlide = pres.addSlide();
      finalSlide.background = { color: "141414" };

      finalSlide.addText("TOP VIVA EXAM QUESTIONS & ANSWERS", {
        x: 0.8,
        y: 0.6,
        w: 11.7,
        h: 0.6,
        fontSize: 22,
        fontFace: "Georgia",
        bold: true,
        italic: true,
        color: "E4E3E0",
      });

      VIVA_QUESTIONS.slice(0, 3).forEach((v, idx) => {
        const itemY = 1.5 + (idx * 1.6);
        finalSlide.addText(`VIVA Q: ${v.question}`, {
          x: 0.8,
          y: itemY,
          w: 11.7,
          h: 0.4,
          fontSize: 12,
          fontFace: "Arial",
          bold: true,
          color: "00FF00",
        });

        finalSlide.addText(`IDEAL ANSWER: ${v.answer}`, {
          x: 0.8,
          y: itemY + 0.35,
          w: 11.7,
          h: 0.6,
          fontSize: 11,
          fontFace: "Arial",
          color: "E4E3E0",
          valign: "top",
        });

        finalSlide.addText(`PRO TIP: ${v.proTip}`, {
          x: 0.8,
          y: itemY + 0.9,
          w: 11.7,
          h: 0.3,
          fontSize: 10,
          fontFace: "Courier New",
          italic: true,
          color: "CCCCCC",
        });
      });

      setPptxStatus("Packaging presentation as file...");
      
      // Try writeFile first. If blocked, fallback to manual Base64 link
      try {
        await pres.writeFile({ fileName: "RDBMS_2505MJCT201_Revision_Masterclass.pptx" });
        setPptxStatus("PowerPoint downloaded successfully!");
      } catch (writeErr: any) {
        console.warn("writeFile failed, trying base64 fallback:", writeErr);
        setPptxStatus("Direct export restricted by browser sandboxing. Generating custom download link...");
        
        const dataUri = await pres.write("base64");
        const base64Content = typeof dataUri === "string" ? dataUri : "";
        
        if (base64Content) {
          setPptxBase64(base64Content);
          
          const link = document.createElement("a");
          link.href = "data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64," + base64Content;
          link.download = "RDBMS_2505MJCT201_Revision_Masterclass.pptx";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          
          setPptxStatus("Generated successfully using sandbox fallback!");
        } else {
          throw writeErr || new Error("Failed to export raw slides data.");
        }
      }
    } catch (error: any) {
      console.error("Error generating PowerPoint:", error);
      setPptxError(error?.message || String(error));
      setPptxStatus("Failed to generate presentation.");
    } finally {
      setIsGeneratingPPTX(false);
    }
  };

  // AI Elaboration & Notes State
  const [aiElaboration, setAiElaboration] = useState<string>("");
  const [isElaborating, setIsElaborating] = useState<boolean>(false);
  const [elaborationTopic, setElaborationTopic] = useState<string>("");

  // AI Custom Tutor State
  const [customQuestion, setCustomQuestion] = useState<string>("");
  const [customAiResponse, setCustomAiResponse] = useState<string>("");
  const [isCustomLoading, setIsCustomLoading] = useState<boolean>(false);

  // AI Mock Exam Evaluator State
  const [selectedQuestion, setSelectedQuestion] = useState<ExpectedQuestion>(EXPECTED_QUESTIONS[0]);
  const [userAnswer, setUserAnswer] = useState<string>("");
  const [evaluationResult, setEvaluationResult] = useState<{
    score: number;
    verdict: string;
    strengths: string[];
    gaps: string[];
    modelAnswer: string;
  } | null>(null);
  const [isEvaluating, setIsEvaluating] = useState<boolean>(false);

  // Selected Viva state
  const [revealedVivaIds, setRevealedVivaIds] = useState<Set<string>>(new Set());

  // Interactive PL/pgSQL Simulator Engine State
  const [simScenario, setSimScenario] = useState<"audit-trigger" | "cursor-loop" | "type-decl">("type-decl");
  const [simStep, setSimStep] = useState<number>(0);
  const [simLogs, setSimLogs] = useState<string[]>(["[Simulator Initialized] Select a script above and click Step to trace execution."]);
  const [simVariables, setSimVariables] = useState<Record<string, { type: string; value: any }>>({});
  const [simCursorState, setSimCursorState] = useState<{ status: string; rowCount: number; found: boolean }>({
    status: "CLOSED",
    rowCount: 0,
    found: false,
  });
  const [simDbEmployees, setSimDbEmployees] = useState<Array<{ id: number; name: string; dept_id: number; salary: number }>>([
    { id: 101, name: "Alice Smith", dept_id: 1, salary: 75000 },
    { id: 102, name: "Bob Jones", dept_id: 2, salary: 62000 },
    { id: 103, name: "Charlie Roy", dept_id: 1, salary: 89000 },
  ]);
  const [simDbAuditLogs, setSimDbAuditLogs] = useState<Array<{ emp_id: number; action: string; timestamp: string }>>([]);

  // Auto-select topic when unit changes
  useEffect(() => {
    const unit = STUDY_UNITS.find(u => u.id === selectedUnitId);
    if (unit && unit.topics.length > 0) {
      // Find if current selected topic belongs to this unit; if not, change it
      const hasTopic = unit.topics.some(t => t.id === selectedTopicId);
      if (!hasTopic) {
        setSelectedTopicId(unit.topics[0].id);
      }
    }
  }, [selectedUnitId]);

  // Find active unit & topic objects
  const activeUnit = STUDY_UNITS.find(u => u.id === selectedUnitId) || STUDY_UNITS[0];
  const activeTopic = activeUnit.topics.find(t => t.id === selectedTopicId) || activeUnit.topics[0];

  // Search filtered topics
  const allFilteredTopics = STUDY_UNITS.flatMap(unit => 
    unit.topics.filter(topic => 
      topic.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      topic.definition.toLowerCase().includes(searchQuery.toLowerCase()) ||
      (topic.example && topic.example.toLowerCase().includes(searchQuery.toLowerCase()))
    ).map(topic => ({ ...topic, unitId: unit.id, unitTitle: unit.title }))
  );

  // Trigger AI Elaboration / Deep Notes
  const handleAskAIElaborate = async (topicTitle: string) => {
    setIsElaborating(true);
    setElaborationTopic(topicTitle);
    setAiElaboration("");
    try {
      const response = await fetch("/api/tutor/explain", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          topic: topicTitle,
          format: "exam-answer",
          additionalPrompt: "Please include a quick last-minute checklist and expected marks advice."
        })
      });
      const data = await response.json();
      if (data.text) {
        setAiElaboration(data.text);
      } else if (data.error) {
        setAiElaboration(`Error: ${data.error}`);
      }
    } catch (err: any) {
      setAiElaboration(`Failed to connect to the AI tutor service. ${err.message || ""}`);
    } finally {
      setIsElaborating(false);
    }
  };

  // Custom AI Search Ask Anything
  const handleCustomQuestion = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!customQuestion.trim()) return;
    setIsCustomLoading(true);
    setCustomAiResponse("");
    try {
      const response = await fetch("/api/tutor/explain", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          topic: customQuestion,
          format: "revision-note",
          additionalPrompt: "Write from the perspective of an RDBMS professor grading college theory exams."
        })
      });
      const data = await response.json();
      if (data.text) {
        setCustomAiResponse(data.text);
      } else {
        setCustomAiResponse("No response received from AI.");
      }
    } catch (err) {
      setCustomAiResponse("Failed to communicate with AI.");
    } finally {
      setIsCustomLoading(false);
    }
  };

  // Evaluate Student Mock Answer
  const handleEvaluateAnswer = async () => {
    if (!userAnswer.trim()) return;
    setIsEvaluating(true);
    setEvaluationResult(null);
    try {
      const response = await fetch("/api/tutor/evaluate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          question: selectedQuestion.question,
          userAnswer: userAnswer
        })
      });
      const data = await response.json();
      setEvaluationResult(data);
    } catch (err) {
      alert("Failed to evaluate answer. Ensure server is active and Gemini API key is configured.");
    } finally {
      setIsEvaluating(false);
    }
  };

  // Toggle Viva answers
  const toggleVivaReveal = (id: string) => {
    const updated = new Set(revealedVivaIds);
    if (updated.has(id)) {
      updated.delete(id);
    } else {
      updated.add(id);
    }
    setRevealedVivaIds(updated);
  };

  // Reset PL/pgSQL simulator
  const handleResetSimulator = (scenarioType = simScenario) => {
    setSimStep(0);
    setSimDbEmployees([
      { id: 101, name: "Alice Smith", dept_id: 1, salary: 75000 },
      { id: 102, name: "Bob Jones", dept_id: 2, salary: 62000 },
      { id: 103, name: "Charlie Roy", dept_id: 1, salary: 89000 },
    ]);
    setSimDbAuditLogs([]);
    setSimCursorState({ status: "CLOSED", rowCount: 0, found: false });

    if (scenarioType === "type-decl") {
      setSimVariables({
        v_emp_id: { type: "employees.emp_id%TYPE", value: "NULL" },
        v_salary: { type: "NUMERIC", value: "NULL" },
        v_bonus: { type: "NUMERIC", value: "NULL" }
      });
      setSimLogs(["[Scenario: %TYPE Declarations] Click 'Step Forward' to declare, bind, and calculate bonus."]);
    } else if (scenarioType === "cursor-loop") {
      setSimVariables({
        v_curr_id: { type: "INT", value: "NULL" },
        v_curr_salary: { type: "NUMERIC", value: "NULL" }
      });
      setSimLogs(["[Scenario: Explicit Cursors] Click 'Step Forward' to open, fetch row-by-row, and process salary increases."]);
    } else {
      setSimVariables({});
      setSimLogs(["[Scenario: Audit Trigger] Click 'Step Forward' to trigger a BEFORE/AFTER Row level execution."]);
    }
  };

  // Initialize Simulator on Mount or Type swap
  useEffect(() => {
    handleResetSimulator(simScenario);
  }, [simScenario]);

  // Step through simulator sequence
  const handleStepSimulator = () => {
    const nextStep = simStep + 1;
    setSimStep(nextStep);

    if (simScenario === "type-decl") {
      switch (nextStep) {
        case 1:
          setSimVariables(prev => ({
            ...prev,
            v_emp_id: { type: "employees.emp_id%TYPE", value: 101 }
          }));
          setSimLogs(prev => [...prev, "Step 1: Assigned v_emp_id = 101. Bound dynamic %TYPE matching table employees."]);
          break;
        case 2:
          // Find employee 101
          const emp = simDbEmployees.find(e => e.id === 101);
          setSimVariables(prev => ({
            ...prev,
            v_salary: { type: "NUMERIC", value: emp ? emp.salary : 0 }
          }));
          setSimLogs(prev => [
            ...prev, 
            `Step 2: Queried SELECT salary INTO v_salary FROM employees. Found balance: ${emp ? emp.salary : 0}`
          ]);
          break;
        case 3:
          const sal = simVariables.v_salary.value;
          const bonusAmt = Math.round(sal * 0.15);
          setSimVariables(prev => ({
            ...prev,
            v_bonus: { type: "NUMERIC", value: bonusAmt }
          }));
          setSimLogs(prev => [
            ...prev,
            `Step 3: Calculated v_bonus := v_salary * 0.15 -> Got ${bonusAmt}.`
          ]);
          break;
        case 4:
          setSimLogs(prev => [
            ...prev,
            `Step 4: RAISE NOTICE 'Employee 101 gets total bonus: ${simVariables.v_bonus.value}'`,
            "SUCCESS: PL/pgSQL Block completed cleanly without exceptions."
          ]);
          break;
        default:
          setSimLogs(prev => [...prev, "Block execution finished. Click reset to repeat."]);
          break;
      }
    } else if (simScenario === "cursor-loop") {
      switch (nextStep) {
        case 1:
          setSimCursorState({ status: "OPEN", rowCount: 0, found: true });
          setSimLogs(prev => [...prev, "Step 1: OPEN emp_cur. Private context memory allocated. Pointer at row 0."]);
          break;
        case 2:
          // Fetch first employee
          setSimVariables({
            v_curr_id: { type: "INT", value: 101 },
            v_curr_salary: { type: "NUMERIC", value: 75000 }
          });
          setSimCursorState({ status: "FETCHING", rowCount: 1, found: true });
          setSimLogs(prev => [...prev, "Step 2: FETCH emp_cur INTO v_curr_id, v_curr_salary. Row 1 retrieved (Alice Smith)."]);
          break;
        case 3:
          // Update Alice Smith salary by 10%
          setSimDbEmployees(prev => prev.map(e => e.id === 101 ? { ...e, salary: e.salary * 1.1 } : e));
          setSimLogs(prev => [...prev, "Step 3: LOOP action -> UPDATE employees SET salary = salary * 1.1 WHERE id = 101. Salary updated."]);
          break;
        case 4:
          // Fetch second employee
          setSimVariables({
            v_curr_id: { type: "INT", value: 102 },
            v_curr_salary: { type: "NUMERIC", value: 62000 }
          });
          setSimCursorState({ status: "FETCHING", rowCount: 2, found: true });
          setSimLogs(prev => [...prev, "Step 4: FETCH emp_cur. Row 2 retrieved (Bob Jones). Cursor attribute FOUND is TRUE."]);
          break;
        case 5:
          // Update Bob Jones salary by 10%
          setSimDbEmployees(prev => prev.map(e => e.id === 102 ? { ...e, salary: e.salary * 1.1 } : e));
          setSimLogs(prev => [...prev, "Step 5: LOOP action -> UPDATE employees SET salary = salary * 1.1 WHERE id = 102. Salary updated."]);
          break;
        case 6:
          // Fetch third employee
          setSimVariables({
            v_curr_id: { type: "INT", value: 103 },
            v_curr_salary: { type: "NUMERIC", value: 89000 }
          });
          setSimCursorState({ status: "FETCHING", rowCount: 3, found: true });
          setSimLogs(prev => [...prev, "Step 6: FETCH emp_cur. Row 3 retrieved (Charlie Roy). Cursor attribute FOUND is TRUE."]);
          break;
        case 7:
          // Update Charlie Roy
          setSimDbEmployees(prev => prev.map(e => e.id === 103 ? { ...e, salary: e.salary * 1.1 } : e));
          setSimLogs(prev => [...prev, "Step 7: LOOP action -> UPDATE employees SET salary = salary * 1.1 WHERE id = 103."]);
          break;
        case 8:
          // Fetch no more rows
          setSimCursorState({ status: "FETCHING", rowCount: 3, found: false });
          setSimLogs(prev => [
            ...prev,
            "Step 8: FETCH emp_cur -> No more rows found. Attribute FOUND became FALSE. Exited LOOP automatically."
          ]);
          break;
        case 9:
          setSimCursorState({ status: "CLOSED", rowCount: 3, found: false });
          setSimLogs(prev => [
            ...prev,
            "Step 9: CLOSE emp_cur. Memory pointer released.",
            "SUCCESS: All rows successfully updated sequentially using the Cursor."
          ]);
          break;
        default:
          setSimLogs(prev => [...prev, "Cursor trace completed."]);
          break;
      }
    } else if (simScenario === "audit-trigger") {
      switch (nextStep) {
        case 1:
          setSimLogs(prev => [
            ...prev,
            "Step 1: Attempting operation -> INSERT INTO employees(id, name, dept_id, salary) VALUES(104, 'Chetana A.', 1, 95000);"
          ]);
          break;
        case 2:
          setSimLogs(prev => [
            ...prev,
            "Step 2: [BEFORE Trigger Intercepts] Checking NOT NULL and structural parameters. Trigger returns NEW."
          ]);
          break;
        case 3:
          // Add employee to virtual DB
          setSimDbEmployees(prev => [...prev, { id: 104, name: "Chetana A.", dept_id: 1, salary: 95000 }]);
          setSimLogs(prev => [
            ...prev,
            "Step 3: Base table insertion succeeded. Row written to non-volatile table memory."
          ]);
          break;
        case 4:
          // Insert audit log automatically!
          const newLog = {
            emp_id: 104,
            action: "INSERT",
            timestamp: new Date().toLocaleTimeString()
          };
          setSimDbAuditLogs([newLog]);
          setSimLogs(prev => [
            ...prev,
            "Step 4: [AFTER INSERT FOR EACH ROW Trigger Fire] -> Executed log_emp_changes() automatically.",
            `SUCCESS: Inserted automatic audit log record for Emp ID 104. Action: ${newLog.action}`
          ]);
          break;
        default:
          setSimLogs(prev => [...prev, "Trigger execution trace completed."]);
          break;
      }
    }
  };

  // Notsopedia Data Fetch & Interaction Handlers
  const fetchNotes = async () => {
    setIsLoadingNotes(true);
    setNotesError("");
    try {
      const res = await fetch("/api/notes");
      if (res.ok) {
        const data = await res.json();
        setUserNotes(data);
        localStorage.setItem("notsopedia_notes_cache", JSON.stringify(data));
      } else {
        throw new Error("Failed to load community notes");
      }
    } catch (err: any) {
      console.error("Fetch notes error:", err);
      setNotesError("Could not retrieve live notes. Displaying cached notes.");
      const local = localStorage.getItem("notsopedia_notes_cache");
      if (local) {
        try {
          setUserNotes(JSON.parse(local));
        } catch (_) {}
      }
    } finally {
      setIsLoadingNotes(false);
    }
  };

  const handleAddNote = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newNoteTitle.trim() || !newNoteContent.trim() || !newNoteSubjectName.trim() || !newNoteUploaderName.trim()) {
      setNotesError("Please fill out all required fields marked with *");
      return;
    }

    setIsSubmittingNote(true);
    setNotesError("");
    setNotesSuccessMessage("");

    const payload = {
      title: newNoteTitle,
      content: newNoteContent,
      subjectName: newNoteSubjectName,
      subjectCode: newNoteSubjectCode,
      topicName: newNoteTopicName,
      uploaderName: newNoteUploaderName,
      uploaderRole: newNoteUploaderRole,
      uploaderEmail: newNoteUploaderEmail
    };

    try {
      const res = await fetch("/api/notes", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      if (res.ok) {
        const createdNote = await res.json();
        const updated = [createdNote, ...userNotes];
        setUserNotes(updated);
        localStorage.setItem("notsopedia_notes_cache", JSON.stringify(updated));
        
        // Reset form fields
        setNewNoteTitle("");
        setNewNoteContent("");
        setNewNoteTopicName("");
        setNewNoteUploaderName("");
        setNewNoteUploaderEmail("");
        setShowUploadForm(false);
        setNotesSuccessMessage("Success! Your note was uploaded to Notsopedia.");
        
        setTimeout(() => setNotesSuccessMessage(""), 6000);
      } else {
        const errorData = await res.json();
        throw new Error(errorData.error || "Failed to upload note");
      }
    } catch (err: any) {
      console.error("Upload note error:", err);
      // Failover to client-side local state & local storage
      const offlineNote: UserNote = {
        id: "offline-" + Date.now(),
        ...payload,
        uploadedAt: new Date().toISOString(),
        likes: 0
      };
      const updated = [offlineNote, ...userNotes];
      setUserNotes(updated);
      localStorage.setItem("notsopedia_notes_cache", JSON.stringify(updated));
      
      setNewNoteTitle("");
      setNewNoteContent("");
      setNewNoteTopicName("");
      setNewNoteUploaderName("");
      setNewNoteUploaderEmail("");
      setShowUploadForm(false);
      setNotesSuccessMessage("Note saved locally (unreachable server).");
      setTimeout(() => setNotesSuccessMessage(""), 6000);
    } finally {
      setIsSubmittingNote(false);
    }
  };

  const handleLikeNote = async (id: string) => {
    // Optimistic UI updates
    setUserNotes(prev => prev.map(n => n.id === id ? { ...n, likes: (n.likes || 0) + 1 } : n));
    try {
      await fetch(`/api/notes/${id}/like`, { method: "POST" });
    } catch (err) {
      console.warn("Could not save like to server, kept in local session", err);
    }
  };

  const fetchSystemConfig = async () => {
    try {
      const res = await fetch("/api/system/config");
      if (res.ok) {
        const data = await res.json();
        setSystemConfig(data);
      }
    } catch (err) {
      console.warn("Could not retrieve system configuration from backend:", err);
    }
  };

  const saveSystemConfig = async (updatedFields: Partial<typeof systemConfig>) => {
    setIsSavingConfig(true);
    try {
      const res = await fetch("/api/system/config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(updatedFields)
      });
      if (res.ok) {
        const data = await res.json();
        setSystemConfig(data);
        setNotesSuccessMessage("Application configuration updated successfully!");
        setTimeout(() => setNotesSuccessMessage(""), 4000);
      }
    } catch (err) {
      console.error("Error saving system config:", err);
      setNotesError("Failed to save app configuration.");
    } finally {
      setIsSavingConfig(false);
    }
  };

  const handleDeleteNote = async (id: string) => {
    if (!confirm("Are you sure you want to delete this note from Notsopedia?")) return;
    try {
      const res = await fetch(`/api/notes/${id}`, { method: "DELETE" });
      if (res.ok) {
        setUserNotes(prev => prev.filter(n => n.id !== id));
        setNotesSuccessMessage("Note deleted successfully.");
        setTimeout(() => setNotesSuccessMessage(""), 4000);
      } else {
        throw new Error("Failed to delete from server");
      }
    } catch (err) {
      console.error("Error deleting note:", err);
      // Local state fallback
      setUserNotes(prev => prev.filter(n => n.id !== id));
      setNotesSuccessMessage("Note deleted locally.");
      setTimeout(() => setNotesSuccessMessage(""), 4000);
    }
  };

  const handleEditNoteSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingNote) return;
    try {
      const res = await fetch(`/api/notes/${editingNote.id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(editingNote)
      });
      if (res.ok) {
        const updated = await res.json();
        setUserNotes(prev => prev.map(n => n.id === updated.id ? updated : n));
        setEditingNote(null);
        setNotesSuccessMessage("Note updated successfully!");
        setTimeout(() => setNotesSuccessMessage(""), 4000);
      } else {
        throw new Error("Failed to save edits to server");
      }
    } catch (err) {
      console.error("Error editing note:", err);
      // Local state fallback
      setUserNotes(prev => prev.map(n => n.id === editingNote.id ? editingNote : n));
      setEditingNote(null);
      setNotesSuccessMessage("Note changes applied locally.");
      setTimeout(() => setNotesSuccessMessage(""), 4000);
    }
  };

  const handleResetDatabase = async () => {
    if (!confirm("Reset all notes to factory defaults? This will erase custom additions!")) return;
    try {
      const res = await fetch("/api/notes/reset", { method: "POST" });
      if (res.ok) {
        const data = await res.json();
        setUserNotes(data);
        setNotesSuccessMessage("Database reset to course defaults.");
        setTimeout(() => setNotesSuccessMessage(""), 4000);
      }
    } catch (err) {
      console.error("Error resetting db:", err);
    }
  };

  useEffect(() => {
    fetchNotes();
    fetchSystemConfig();
  }, []);

  useEffect(() => {
    if (activeTab === "notesopedia") {
      fetchNotes();
      fetchSystemConfig();
    }
  }, [activeTab]);

  return (
    <div className="flex flex-col h-screen overflow-hidden text-sm bg-[#E4E3E0] text-[#141414] font-sans">
      
      {/* PPTX Generation Notification Banner */}
      {(pptxStatus || pptxError) && (
        <div className={`px-6 py-2.5 font-mono text-xs flex flex-wrap items-center justify-between gap-3 border-b-2 border-black ${pptxError ? "bg-rose-100 text-rose-950 animate-fadeIn" : "bg-emerald-50 text-emerald-950 animate-fadeIn"}`}>
          <div className="flex items-center space-x-2">
            <span className={`w-2.5 h-2.5 rounded-full inline-block ${pptxError ? "bg-rose-600" : isGeneratingPPTX ? "bg-emerald-500 animate-ping" : "bg-emerald-500"}`}></span>
            <span className="font-bold uppercase">PowerPoint:</span>
            <span>{pptxStatus}</span>
            {pptxError && <span className="text-rose-700 ml-1">({pptxError})</span>}
          </div>
          
          <div className="flex items-center space-x-3">
            {pptxBase64 && (
              <a
                href={`data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${pptxBase64}`}
                download="RDBMS_2505MJCT201_Revision_Masterclass.pptx"
                className="px-3 py-1 bg-black text-white uppercase font-bold text-[10px] tracking-wider rounded border border-black hover:bg-neutral-800 transition-all shadow-[1px_1px_0px_rgba(0,0,0,1)] active:translate-x-[0.5px] active:translate-y-[0.5px]"
              >
                💾 Direct Link Fallback Download
              </a>
            )}
            
            {pptxError && (
              <span className="text-[10px] opacity-85 bg-white/50 px-2 py-0.5 rounded border border-black/10">
                💡 Tip: If blocked by browser sandbox inside the preview, click "Open in New Tab" at the top right of the screen!
              </span>
            )}
            
            <button 
              onClick={() => { setPptxStatus(""); setPptxError(""); }}
              className="text-black font-bold px-1.5 hover:bg-black/10 rounded"
              title="Dismiss"
            >
              ✕
            </button>
          </div>
        </div>
      )}

      {/* Global Admin Announcement Banner */}
      {systemConfig && systemConfig.announcementActive && systemConfig.announcement && (
        <div className="bg-[#10b981] text-black px-6 py-2 font-mono text-[11px] tracking-wide font-semibold flex items-center justify-between border-b border-black select-none">
          <div className="flex items-center space-x-2 overflow-hidden truncate">
            <span className="bg-black text-[#10b981] text-[9px] px-1.5 py-0.5 rounded uppercase font-extrabold tracking-wider shrink-0 animate-pulse">Broadcast</span>
            <span className="truncate">{systemConfig.announcement}</span>
          </div>
          <button 
            onClick={() => setSystemConfig(prev => ({ ...prev, announcementActive: false }))} 
            className="text-black font-black hover:bg-black/10 px-1 rounded ml-3 shrink-0"
            title="Dismiss Announcement"
          >
            ✕
          </button>
        </div>
      )}
      
      {/* HEADER SECTION */}
      <header className="flex items-center justify-between px-6 py-4 bg-[#E4E3E0] tech-border-b">
        <div className="flex items-center space-x-3">
          <div className="flex items-center justify-center w-8 h-8 font-mono text-base font-bold text-[#E4E3E0] bg-black rounded-sm">
            N
          </div>
          <div>
            <h1 className="text-xl font-bold font-serif italic tracking-tight">
              Notsopedia
            </h1>
            <p className="text-[10px] font-mono tracking-widest uppercase opacity-75">
              COLLABORATIVE STUDY HUBS • RDBMS STUDY & EXAM MASTERCLASS
            </p>
          </div>
        </div>

        {/* Navigation Tabs - High Contrast Buttons */}
        <div className="flex items-center space-x-1 font-mono text-xs">
          <button
            onClick={() => setActiveTab("syllabus")}
            className={`px-3 py-1.5 border border-black uppercase font-semibold transition ${
              activeTab === "syllabus" ? "bg-black text-white" : "hover:bg-black/5"
            }`}
          >
            Syllabus Core Notes
          </button>
          <button
            onClick={() => setActiveTab("simulator")}
            className={`px-3 py-1.5 border border-black uppercase font-semibold transition ${
              activeTab === "simulator" ? "bg-black text-white" : "hover:bg-black/5"
            }`}
          >
            PL/pgSQL Code Simulator
          </button>
          <button
            onClick={() => setActiveTab("mock-exam")}
            className={`px-3 py-1.5 border border-black uppercase font-semibold transition ${
              activeTab === "mock-exam" ? "bg-black text-white" : "hover:bg-black/5"
            }`}
          >
            Mock Grader (AI)
          </button>
          <button
            onClick={() => setActiveTab("viva")}
            className={`px-3 py-1.5 border border-black uppercase font-semibold transition ${
              activeTab === "viva" ? "bg-black text-white" : "hover:bg-black/5"
            }`}
          >
            Viva Voce & Tricks
          </button>
          <button
            onClick={() => setActiveTab("notesopedia")}
            className={`px-3 py-1.5 border border-black uppercase font-bold text-xs transition flex items-center space-x-1.5 ${
              activeTab === "notesopedia" ? "bg-[#10b981] text-white" : "bg-emerald-50 hover:bg-emerald-100"
            }`}
          >
            <span>📚 Notsopedia Hub</span>
            {userNotes.length > 0 && (
              <span className="bg-black text-white px-1 py-0.5 text-[9px] rounded-full font-bold">
                {userNotes.length}
              </span>
            )}
          </button>
        </div>

        {/* Syllabus / Exam Specs Badge */}
        <div className="flex items-center space-x-4">
          <button
            onClick={handleDownloadPPTX}
            disabled={isGeneratingPPTX}
            id="download-pptx-button"
            className={`px-4 py-2 border-2 border-black uppercase font-bold text-xs tracking-wider transition-all flex items-center space-x-2 shadow-[2px_2px_0px_0px_rgba(0,0,0,1)] hover:shadow-[4px_4px_0px_0px_rgba(0,0,0,1)] active:translate-x-[2px] active:translate-y-[2px] active:shadow-[0px_0px_0px_0px_rgba(0,0,0,1)] ${
              isGeneratingPPTX 
                ? "bg-black/10 text-black/40 cursor-not-allowed" 
                : "bg-emerald-400 text-[#141414] hover:bg-emerald-500"
            }`}
            title="Download full course revision PowerPoint presentation"
          >
            <FileText className={`h-4 w-4 ${isGeneratingPPTX ? "animate-spin" : ""}`} />
            <span>{isGeneratingPPTX ? "Generating..." : "Download PPTX Presentation"}</span>
          </button>

          <div className="text-right hidden md:block">
            <div className="font-mono text-[10px] opacity-75">TIME: 2.0 HOURS</div>
            <div className="font-mono text-[10px] font-bold">TOTAL: 30 MARKS</div>
          </div>
          <div className="flex items-center justify-center w-10 h-10 border-2 border-black rounded-full font-bold font-serif text-lg bg-[#141414] text-[#E4E3E0]">
            30
          </div>
        </div>
      </header>

      {/* CORE SPLIT SCREEN GRID */}
      <div className="flex flex-1 overflow-hidden">
        
        {/* SIDEBAR LEFT PANEL - UNIT & SEARCH (Only visible in Syllabus & Viva pages) */}
        {(activeTab === "syllabus" || activeTab === "viva") && (
          <aside className="w-80 flex flex-col bg-black/3 tech-border-r h-full overflow-y-auto">
            
            {/* Search Syllabus */}
            <div className="p-4 border-b border-black">
              <label className="block text-[10px] font-mono uppercase tracking-wider mb-1.5 font-bold">
                Search Syllabus & Notes
              </label>
              <div className="relative">
                <Search className="absolute left-2.5 top-2.5 h-4 w-4 opacity-50" />
                <input
                  type="text"
                  placeholder="e.g. ACID, Cursor, 1NF..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full pl-9 pr-3 py-2 bg-[#E4E3E0] border border-black rounded-none text-xs focus:outline-none focus:ring-1 focus:ring-black placeholder-black/50"
                />
              </div>
              {searchQuery && (
                <button 
                  onClick={() => setSearchQuery("")}
                  className="mt-1.5 text-[10px] font-mono underline hover:text-black/70"
                >
                  Clear search filters
                </button>
              )}
            </div>

            {/* If there's an active search query, show global results */}
            {searchQuery ? (
              <div className="flex-1 flex flex-col">
                <div className="bg-black text-white px-4 py-2 text-[10px] font-mono uppercase tracking-widest">
                  Found ({allFilteredTopics.length}) Matches
                </div>
                <div className="flex-1 overflow-y-auto divide-y divide-black/10">
                  {allFilteredTopics.map((topic) => (
                    <button
                      key={topic.id}
                      onClick={() => {
                        setSelectedUnitId(topic.unitId);
                        setSelectedTopicId(topic.id);
                        if (activeTab !== "syllabus") {
                          setActiveTab("syllabus");
                        }
                      }}
                      className={`w-full text-left p-4 hover:bg-black/5 transition flex flex-col ${
                        selectedTopicId === topic.id ? "bg-black/10" : ""
                      }`}
                    >
                      <span className="text-[10px] font-mono uppercase tracking-wider text-black/50">
                        {topic.unitTitle.split(":")[0]}
                      </span>
                      <span className="font-serif font-bold italic text-sm mt-0.5">
                        {topic.title}
                      </span>
                      <p className="text-xs text-black/70 mt-1 line-clamp-2">
                        {topic.definition}
                      </p>
                    </button>
                  ))}
                  {allFilteredTopics.length === 0 && (
                    <div className="p-4 text-center font-mono text-xs opacity-60">
                      No topics matched your query. Try searching for "2PL" or "Trigger".
                    </div>
                  )}
                </div>
              </div>
            ) : (
              // Otherwise show list of Units & their Topics
              <div className="flex-1 flex flex-col divide-y divide-black/30">
                {STUDY_UNITS.map((unit) => {
                  const isUnitSelected = selectedUnitId === unit.id;
                  return (
                    <div key={unit.id} className="flex flex-col">
                      {/* Unit Header Button */}
                      <button
                        onClick={() => setSelectedUnitId(unit.id)}
                        className={`w-full text-left p-4 transition duration-150 flex flex-col ${
                          isUnitSelected ? "bg-black text-white" : "hover:bg-black/5"
                        }`}
                      >
                        <span className="text-[10px] font-mono tracking-widest uppercase opacity-75">
                          {unit.id.toUpperCase().replace("-", " ")}
                        </span>
                        <h3 className="font-serif font-bold italic text-sm leading-tight mt-1">
                          {unit.title}
                        </h3>
                        <p className={`text-[11px] mt-1 line-clamp-1 ${isUnitSelected ? "text-[#E4E3E0]/70" : "text-black/60"}`}>
                          {unit.description}
                        </p>
                      </button>

                      {/* Topic Submenu (Only active unit expands) */}
                      {isUnitSelected && (
                        <div className="bg-[#E4E3E0]/40 py-1.5 divide-y divide-black/5">
                          {unit.topics.map((topic) => {
                            const isTopicSelected = selectedTopicId === topic.id;
                            return (
                              <button
                                key={topic.id}
                                onClick={() => setSelectedTopicId(topic.id)}
                                className={`w-full text-left px-6 py-2.5 text-xs transition flex items-center justify-between ${
                                  isTopicSelected ? "bg-black/10 font-bold" : "hover:bg-black/5"
                                }`}
                              >
                                <span className="truncate pr-2 font-mono">
                                  ▸ {topic.title}
                                </span>
                                <ChevronRight className="h-3 w-3 flex-shrink-0 opacity-40" />
                              </button>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}

            {/* Quick Stats Block */}
            <div className="p-4 border-t border-black bg-[#141414] text-[#E4E3E0] mt-auto">
              <div className="flex justify-between items-center text-[10px] font-mono uppercase tracking-wider">
                <span>RDBMS Exam Readiness</span>
                <span className="font-bold text-emerald-400">92%</span>
              </div>
              <div className="w-full bg-[#E4E3E0]/20 h-1 mt-1.5 overflow-hidden">
                <div className="bg-emerald-400 h-full" style={{ width: '92%' }}></div>
              </div>
              <p className="text-[10px] mt-1.5 opacity-50 font-serif italic text-right">
                All 6 core modules fully indexed.
              </p>
            </div>
          </aside>
        )}

        {/* MAIN WORKING CANVAS */}
        <main className="flex-1 flex flex-col h-full overflow-y-auto bg-[#E4E3E0] p-6">
          
          {/* TAB 1: SYLLABUS DETAIL VIEWER */}
          {activeTab === "syllabus" && (
            <div className="space-y-6">
              
              {/* Active Header Block */}
              <div className="bg-black text-white p-5 tech-border">
                <span className="text-[10px] font-mono tracking-widest uppercase text-[#E4E3E0]/75">
                  {activeUnit.title.toUpperCase()}
                </span>
                <h2 className="text-2xl font-serif font-bold italic mt-1 text-[#E4E3E0]">
                  {activeTopic.title}
                </h2>
                <p className="text-xs text-[#E4E3E0]/85 mt-2 max-w-3xl leading-relaxed">
                  {activeTopic.definition}
                </p>
              </div>

              {/* Grid Content */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                
                {/* Structural Parameters Card */}
                <div className="bg-transparent border border-black p-5 flex flex-col">
                  <div className="border-b border-black pb-2 mb-4">
                    <span className="font-mono text-[10px] uppercase tracking-wider font-bold">
                      [01] Types, Features & Components
                    </span>
                  </div>
                  {activeTopic.typesOrFeatures && activeTopic.typesOrFeatures.length > 0 ? (
                    <ul className="space-y-3 flex-1">
                      {activeTopic.typesOrFeatures.map((feat, idx) => (
                        <li key={idx} className="flex items-start text-xs">
                          <span className="font-mono mr-2.5 text-black/40 mt-0.5">({idx + 1})</span>
                          <span className="leading-normal">{feat}</span>
                        </li>
                      ))}
                    </ul>
                  ) : (
                    <div className="p-4 text-center italic text-xs text-black/50">
                      Standard PostgreSQL relational database module.
                    </div>
                  )}

                  {/* Advantages section */}
                  <div className="border-t border-black mt-5 pt-4">
                    <span className="font-mono text-[10px] uppercase tracking-wider font-bold block mb-2.5">
                      Key University Advantages
                    </span>
                    {activeTopic.advantages && activeTopic.advantages.length > 0 ? (
                      <ul className="space-y-1.5">
                        {activeTopic.advantages.map((adv, idx) => (
                          <li key={idx} className="text-xs flex items-center space-x-2">
                            <span className="w-1 h-1 bg-black rounded-full"></span>
                            <span>{adv}</span>
                          </li>
                        ))}
                      </ul>
                    ) : (
                      <p className="text-xs text-black/70 italic">
                        Provides complete data isolation, concurrency control, and relational mapping.
                      </p>
                    )}
                  </div>
                </div>

                {/* Practical Syntax & Example Block */}
                <div className="bg-[#141414] text-[#E4E3E0] p-5 flex flex-col font-mono">
                  <div className="border-b border-white/20 pb-2 mb-3 flex justify-between items-center">
                    <span className="text-[10px] uppercase tracking-wider text-emerald-400 font-bold">
                      [02] Exam Code Blueprint
                    </span>
                    <span className="text-[9px] bg-white/10 px-1.5 py-0.5">PostgreSQL / PL/pgSQL</span>
                  </div>

                  {activeTopic.syntax && (
                    <div className="mb-4">
                      <span className="text-[10px] text-white/50 block mb-1">GENERAL SYNTAX:</span>
                      <pre className="text-xs bg-black/50 p-2 border border-white/10 overflow-x-auto text-[#E4E3E0] max-h-40">
                        {activeTopic.syntax}
                      </pre>
                    </div>
                  )}

                  <div className="flex-1 flex flex-col justify-between">
                    <div>
                      <span className="text-[10px] text-white/50 block mb-1">COMPLETE WORKING EXAMPLE:</span>
                      <pre className="text-xs bg-black/60 p-3 rounded-sm text-emerald-300 overflow-x-auto max-h-60 leading-relaxed font-mono">
                        {activeTopic.example || `-- Sample Query\nSELECT * FROM ${activeTopic.id.replace("-", "_")};`}
                      </pre>
                    </div>

                    <div className="mt-4 pt-3 border-t border-white/10 text-[10px] text-white/60 flex items-center justify-between">
                      <span>✓ Syntactically validated on PG-15</span>
                      {activeTopic.example && (
                        <button 
                          onClick={() => {
                            if (activeTopic.id.includes("cursor")) setSimScenario("cursor-loop");
                            else if (activeTopic.id.includes("trigger")) setSimScenario("audit-trigger");
                            else setSimScenario("type-decl");
                            setActiveTab("simulator");
                          }}
                          className="px-2 py-1 bg-emerald-500 hover:bg-emerald-600 text-black font-semibold uppercase tracking-wider text-[9px] flex items-center space-x-1"
                        >
                          <Play className="h-2.5 w-2.5" />
                          <span>Load in Simulator</span>
                        </button>
                      )}
                    </div>
                  </div>
                </div>

              </div>

              {/* DYNAMIC AI EXTENSION AND NOTE ELABORATOR */}
              <div className="border border-black bg-white/30 p-5 mt-6">
                <div className="flex flex-col md:flex-row items-start md:items-center justify-between border-b border-black pb-3 mb-4">
                  <div>
                    <h3 className="font-serif font-bold italic text-base flex items-center space-x-2">
                      <Brain className="h-5 w-5 text-black" />
                      <span>AI Professor: Deep Lecture Explainer</span>
                    </h3>
                    <p className="text-xs text-black/70">
                      Need a full 15-page essay breakdown? Let Gemini generate structured university notes with extra examples.
                    </p>
                  </div>
                  <button
                    onClick={() => handleAskAIElaborate(activeTopic.title)}
                    disabled={isElaborating}
                    className="mt-3 md:mt-0 font-mono text-xs px-4 py-2 border border-black bg-black text-white hover:bg-black/90 font-semibold tracking-wider uppercase flex items-center space-x-2 disabled:opacity-50"
                  >
                    {isElaborating ? (
                      <>
                        <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                        <span>Generating Notes...</span>
                      </>
                    ) : (
                      <>
                        <Sparkles className="h-3.5 w-3.5" />
                        <span>Elaborate Topic with AI</span>
                      </>
                    )}
                  </button>
                </div>

                {/* Show notes if loading or filled */}
                {(isElaborating || aiElaboration) && (
                  <div className="bg-white p-5 border border-black font-mono text-xs leading-relaxed max-h-[500px] overflow-y-auto">
                    {isElaborating ? (
                      <div className="space-y-3">
                        <div className="flex items-center space-x-2 text-black/60">
                          <RefreshCw className="h-4 w-4 animate-spin" />
                          <span>Gemini is compiling syllabus schemas, drafting functional dependencies, and writing complete PL/pgSQL solutions...</span>
                        </div>
                        <div className="h-4 bg-black/5 animate-pulse rounded w-1/3"></div>
                        <div className="h-4 bg-black/5 animate-pulse rounded w-3/4"></div>
                        <div className="h-4 bg-black/5 animate-pulse rounded w-5/6"></div>
                      </div>
                    ) : (
                      <div className="prose prose-sm max-w-none text-[#141414]">
                        <div className="flex justify-between items-center mb-4 bg-black/5 p-2 border-b border-black">
                          <span className="font-bold uppercase tracking-wider text-[10px]">LECTURER NOTE FOR: {elaborationTopic.toUpperCase()}</span>
                          <span className="text-[9px] bg-black text-white px-1.5 py-0.5">Gemini 3.5 Flash</span>
                        </div>
                        <div className="whitespace-pre-wrap font-sans text-sm leading-relaxed">
                          {aiElaboration}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* CUSTOM AI QUESTION BAR */}
              <div className="border border-black bg-black text-[#E4E3E0] p-6">
                <span className="font-mono text-[9px] uppercase tracking-widest text-emerald-400 block mb-1">
                  [Interactive Ask-Anything Panel]
                </span>
                <h3 className="font-serif font-bold italic text-lg mb-2 text-white">
                  Have a specific RDBMS question in mind?
                </h3>
                <p className="text-xs text-white/70 mb-4 max-w-2xl">
                  Ask about BCNF vs 3NF, multi-version concurrency control (MVCC), or write custom trigger code. Our AI RDBMS professor will write a pristine academic answer immediately.
                </p>

                <form onSubmit={handleCustomQuestion} className="space-y-4">
                  <div className="flex items-center space-x-2">
                    <input
                      type="text"
                      placeholder="e.g. Explain View Serializability vs Conflict Serializability with a schedule table"
                      value={customQuestion}
                      onChange={(e) => setCustomQuestion(e.target.value)}
                      className="flex-1 bg-[#E4E3E0] text-black border border-white px-3 py-2.5 text-xs rounded-none focus:outline-none focus:ring-1 focus:ring-white placeholder-black/50 font-mono"
                    />
                    <button
                      type="submit"
                      disabled={isCustomLoading || !customQuestion.trim()}
                      className="px-5 py-2.5 bg-emerald-400 text-black font-mono text-xs uppercase font-bold tracking-wider hover:bg-emerald-500 disabled:opacity-50 shrink-0"
                    >
                      {isCustomLoading ? "Thinking..." : "Query AI"}
                    </button>
                  </div>
                </form>

                {/* Response Area */}
                {(isCustomLoading || customAiResponse) && (
                  <div className="mt-4 bg-[#E4E3E0] text-[#141414] p-5 border border-white font-sans text-xs leading-relaxed max-h-96 overflow-y-auto">
                    {isCustomLoading ? (
                      <div className="flex items-center space-x-3 text-black font-mono">
                        <RefreshCw className="h-4 w-4 animate-spin" />
                        <span>Searching academic material and compiling response...</span>
                      </div>
                    ) : (
                      <div className="prose prose-sm max-w-none">
                        <div className="flex justify-between items-center mb-3 pb-2 border-b border-black font-mono text-[10px] tracking-wide uppercase">
                          <span>Custom Tutoring Answer</span>
                          <span>UTC Response</span>
                        </div>
                        <div className="whitespace-pre-wrap font-sans text-sm leading-relaxed">
                          {customAiResponse}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>

            </div>
          )}

          {/* TAB 2: PL/pgSQL CODE RUNNER & SIMULATOR */}
          {activeTab === "simulator" && (
            !systemConfig.enableSimulator && !isAdminMode ? (
              <div className="border-2 border-black bg-white p-8 text-center space-y-4 max-w-2xl mx-auto shadow-[4px_4px_0px_rgba(0,0,0,1)] my-12 font-mono">
                <Lock className="h-12 w-12 text-rose-600 mx-auto animate-bounce" />
                <h3 className="text-base font-black uppercase text-rose-950">
                  ⚠️ Procedural Simulator Locked
                </h3>
                <p className="text-xs text-black/80 leading-relaxed max-w-md mx-auto">
                  The PL/pgSQL Execution and Step-by-Step procedural runtime simulator is currently deactivated for system maintenance by the administrator.
                </p>
                <div className="bg-neutral-50 p-3 border border-neutral-300 text-[11px] leading-tight text-neutral-600">
                  Students may continue accessing curriculum syllabus guides, mock tests, and the Notsopedia Peer-to-Peer shared course directory.
                </div>
              </div>
            ) : (
              <div className="space-y-6">
                
                <div className="bg-black text-white p-5 tech-border">
                  <span className="text-[10px] font-mono tracking-widest uppercase text-[#E4E3E0]/75">
                    PL/pgSQL Execution Simulator
                  </span>
                  <h2 className="text-xl font-serif font-bold italic mt-1 text-[#E4E3E0]">
                    Procedural Runtime Inspector
                  </h2>
                  <p className="text-xs text-[#E4E3E0]/80 mt-1.5 leading-relaxed">
                    Understand how procedural blocks run server-side. Track values, cursors, and triggers through step-by-step memory updates without compiling local databases.
                  </p>
                </div>

              {/* Scenario selector */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
                <button
                  onClick={() => setSimScenario("type-decl")}
                  className={`p-3 border border-black font-mono text-xs text-left uppercase font-bold flex items-center justify-between ${
                    simScenario === "type-decl" ? "bg-black text-[#E4E3E0]" : "bg-transparent hover:bg-black/5"
                  }`}
                >
                  <span>1. %TYPE Declaration</span>
                  <span className="text-[9px] px-1 bg-black text-white py-0.5 font-normal">Active</span>
                </button>
                <button
                  onClick={() => setSimScenario("cursor-loop")}
                  className={`p-3 border border-black font-mono text-xs text-left uppercase font-bold flex items-center justify-between ${
                    simScenario === "cursor-loop" ? "bg-black text-[#E4E3E0]" : "bg-transparent hover:bg-black/5"
                  }`}
                >
                  <span>2. Explicit Cursor loop</span>
                  <span className="text-[9px] px-1 bg-black text-white py-0.5 font-normal">Active</span>
                </button>
                <button
                  onClick={() => setSimScenario("audit-trigger")}
                  className={`p-3 border border-black font-mono text-xs text-left uppercase font-bold flex items-center justify-between ${
                    simScenario === "audit-trigger" ? "bg-black text-[#E4E3E0]" : "bg-transparent hover:bg-black/5"
                  }`}
                >
                  <span>3. Audit Log Trigger</span>
                  <span className="text-[9px] px-1 bg-black text-white py-0.5 font-normal">Active</span>
                </button>
              </div>

              {/* Simulator Main Grid */}
              <div className="grid grid-cols-1 xl:grid-cols-12 gap-6">
                
                {/* Script Code Viewer */}
                <div className="xl:col-span-4 bg-[#141414] text-[#E4E3E0] p-4 flex flex-col justify-between tech-border">
                  <div>
                    <div className="border-b border-white/20 pb-2 mb-3 flex justify-between items-center">
                      <span className="font-mono text-[10px] text-emerald-400 uppercase font-bold">
                        EXECUTION BLOCK
                      </span>
                      <span className="text-[10px] font-mono opacity-50">Step: {simStep}</span>
                    </div>

                    <pre className="font-mono text-xs bg-black/60 p-3 overflow-x-auto leading-normal">
                      {simScenario === "type-decl" && (
`-- Calculates employee 101 bonus
DECLARE
  v_emp_id employees.emp_id%TYPE := 101; -- Line 2
  v_salary NUMERIC;                      -- Line 3
  v_bonus NUMERIC;                       -- Line 4
BEGIN
  -- Query balance from base table
  SELECT salary INTO v_salary            -- Line 7
  FROM employees WHERE emp_id = v_emp_id;

  v_bonus := v_salary * 0.15;            -- Line 10
  
  RAISE NOTICE 'Bonus is %', v_bonus;    -- Line 12
END;`
                      )}

                      {simScenario === "cursor-loop" && (
`-- Iterates and increases salaries
DECLARE
  v_curr_id INT;
  v_curr_salary NUMERIC;
  emp_cur CURSOR FOR SELECT emp_id, salary FROM employees;
BEGIN
  OPEN emp_cur;                          -- Step 1
  LOOP
    FETCH emp_cur INTO v_curr_id, v_curr_salary;
    EXIT WHEN NOT FOUND;                 -- Checks status
    
    UPDATE employees SET salary = v_curr_salary * 1.1 
    WHERE emp_id = v_curr_id;
  END LOOP;
  CLOSE emp_cur;
END;`
                      )}

                      {simScenario === "audit-trigger" && (
`-- Automatically logs changes
CREATE TRIGGER emp_audit_trg
AFTER INSERT ON employees
FOR EACH ROW EXECUTE FUNCTION log_emp_changes();

-- Triggered Statement:
INSERT INTO employees(id, name, dept_id, salary)
VALUES(104, 'Chetana A.', 1, 95000);`
                      )}
                    </pre>
                  </div>

                  <div className="mt-4 pt-3 border-t border-white/10 flex items-center justify-between font-mono text-[11px]">
                    <button
                      onClick={() => handleResetSimulator()}
                      className="px-3 py-1.5 border border-[#E4E3E0] hover:bg-white hover:text-black transition uppercase font-bold"
                    >
                      Reset Sim
                    </button>
                    <button
                      onClick={handleStepSimulator}
                      className="px-4 py-1.5 bg-emerald-400 text-black font-bold uppercase hover:bg-emerald-500 transition flex items-center space-x-1"
                    >
                      <Play className="h-3 w-3" />
                      <span>Step Forward ▸</span>
                    </button>
                  </div>
                </div>

                {/* State Tables & Memory Symbol Table */}
                <div className="xl:col-span-8 space-y-6">
                  
                  {/* Symbol table and logs */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    
                    {/* Simulated Variables */}
                    <div className="border border-black p-4 flex flex-col bg-white bg-opacity-20">
                      <div className="border-b border-black pb-1.5 mb-2.5">
                        <span className="font-mono text-[10px] uppercase tracking-wider font-bold">
                          [REGISTRY] Variable Symbol Table
                        </span>
                      </div>
                      <div className="flex-1 overflow-y-auto max-h-48 font-mono text-xs">
                        {Object.keys(simVariables).length > 0 ? (
                          <table className="w-full text-left">
                            <thead>
                              <tr className="border-b border-black text-[10px] text-black/50">
                                <th className="pb-1">Variable</th>
                                <th className="pb-1">Data Type</th>
                                <th className="pb-1 text-right">Current Value</th>
                              </tr>
                            </thead>
                            <tbody>
                              {Object.entries(simVariables).map(([name, detail]: [string, any]) => (
                                <tr key={name} className="border-b border-black/5">
                                  <td className="py-1.5 font-bold text-indigo-950">{name}</td>
                                  <td className="py-1.5 text-[11px] opacity-70">{detail.type}</td>
                                  <td className="py-1.5 text-right font-bold text-emerald-800">{String(detail.value)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        ) : (
                          <p className="text-xs text-black/50 italic py-4">No active variables in memory.</p>
                        )}

                        {/* Cursor attributes */}
                        {simScenario === "cursor-loop" && (
                          <div className="mt-4 pt-3 border-t border-black/20">
                            <span className="text-[10px] uppercase font-bold opacity-60 block mb-1">Cursor Attributes:</span>
                            <div className="grid grid-cols-3 gap-1.5 text-center text-[10px]">
                              <div className="border border-black/30 p-1 bg-black/5">
                                <div className="opacity-60">STATUS</div>
                                <div className="font-bold">{simCursorState.status}</div>
                              </div>
                              <div className="border border-black/30 p-1 bg-black/5">
                                <div className="opacity-60">%ROWCOUNT</div>
                                <div className="font-bold">{simCursorState.rowCount}</div>
                              </div>
                              <div className="border border-black/30 p-1 bg-black/5">
                                <div className="opacity-60">%FOUND</div>
                                <div className="font-bold text-emerald-800">{simCursorState.found ? "TRUE" : "FALSE"}</div>
                              </div>
                            </div>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Console Logs */}
                    <div className="border border-black bg-black text-emerald-400 p-4 flex flex-col justify-between">
                      <div className="border-b border-white/20 pb-1.5 mb-2">
                        <span className="font-mono text-[10px] uppercase text-white/50 block font-bold">
                          [CONSOLE] RDBMS ENGINE OUTPUT
                        </span>
                      </div>
                      <div className="flex-1 overflow-y-auto max-h-48 font-mono text-[11px] space-y-1 scrollbar-thin scrollbar-thumb-white">
                        {simLogs.map((log, index) => (
                          <div key={index} className="leading-relaxed">
                            <span className="text-white/40">&gt; </span>
                            <span>{log}</span>
                          </div>
                        ))}
                      </div>
                      <div className="mt-2 text-[9px] text-white/50 text-right">
                        Simulating PostgreSQL v15.4 Memory Stack
                      </div>
                    </div>

                  </div>

                  {/* Relational Tables View */}
                  <div className="border border-black p-4 bg-white/40">
                    <div className="border-b border-black pb-2 mb-3 flex justify-between items-center">
                      <span className="font-mono text-[10px] uppercase tracking-wider font-bold">
                        [STORAGE SCHEMA] ACTIVE RELATIONAL TABLES
                      </span>
                      <span className="text-[10px] font-mono text-indigo-950">Table: employees</span>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full text-left font-mono text-xs">
                        <thead>
                          <tr className="border-b-2 border-black text-black/60 text-[10px]">
                            <th className="pb-1.5">emp_id (PK)</th>
                            <th className="pb-1.5">emp_name</th>
                            <th className="pb-1.5">dept_id (FK)</th>
                            <th className="pb-1.5 text-right">salary (NUMERIC)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-black/10">
                          {simDbEmployees.map((emp) => (
                            <tr key={emp.id} className="hover:bg-black/5">
                              <td className="py-2 font-bold">{emp.id}</td>
                              <td className="py-2">{emp.name}</td>
                              <td className="py-2">{emp.dept_id}</td>
                              <td className="py-2 text-right font-semibold text-indigo-900">${emp.salary.toLocaleString()}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {/* Audit logs table if triggered scenario */}
                    {simScenario === "audit-trigger" && (
                      <div className="mt-4 pt-4 border-t border-black/20">
                        <div className="font-mono text-[10px] font-bold text-red-900 mb-1.5 uppercase">
                          Auditing Table: employee_audit
                        </div>
                        <div className="overflow-x-auto">
                          <table className="w-full text-left font-mono text-xs">
                            <thead>
                              <tr className="border-b border-black text-black/60 text-[10px]">
                                <th className="pb-1">emp_id</th>
                                <th className="pb-1">action_type</th>
                                <th className="pb-1 text-right">logged_at</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-black/10">
                              {simDbAuditLogs.map((log, idx) => (
                                <tr key={idx} className="bg-red-50">
                                  <td className="py-1 font-bold">{log.emp_id}</td>
                                  <td className="py-1 text-red-700 font-semibold">{log.action}</td>
                                  <td className="py-1 text-right text-black/60">{log.timestamp}</td>
                                </tr>
                              ))}
                              {simDbAuditLogs.length === 0 && (
                                <tr>
                                  <td colSpan={3} className="py-2 text-center italic text-black/40">
                                    No audit logs captured. Advance simulator steps to insert a record.
                                  </td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                  </div>

                </div>

              </div>

            </div>
            )
          )}

          {/* TAB 3: MOCK EXAM GRADER (AI EVALUATOR) */}
          {activeTab === "mock-exam" && (
            <div className="space-y-6">
              
              <div className="bg-black text-white p-5 tech-border">
                <span className="text-[10px] font-mono tracking-widest uppercase text-[#E4E3E0]/75">
                  University Exam Simulator
                </span>
                <h2 className="text-xl font-serif font-bold italic mt-1 text-[#E4E3E0]">
                  Mock Grader & AI Evaluator
                </h2>
                <p className="text-xs text-[#E4E3E0]/80 mt-1.5 leading-relaxed">
                  Select a highly expected 5-mark university question, write down your theoretical answer in the typewriter box, and get an immediate, detailed academic evaluation with a strict marks breakdown.
                </p>
              </div>

              {/* Questions List */}
              <div className="grid grid-cols-1 md:grid-cols-5 gap-2 font-mono">
                {EXPECTED_QUESTIONS.map((q, idx) => (
                  <button
                    key={q.id}
                    onClick={() => {
                      setSelectedQuestion(q);
                      setEvaluationResult(null);
                      setUserAnswer("");
                    }}
                    className={`p-3 border border-black text-left text-xs uppercase font-bold flex flex-col justify-between ${
                      selectedQuestion.id === q.id ? "bg-black text-[#E4E3E0]" : "bg-transparent hover:bg-black/5"
                    }`}
                  >
                    <span>Question {idx + 1}</span>
                    <span className="text-[9px] opacity-75 mt-1.5 font-normal line-clamp-1">{q.question}</span>
                  </button>
                ))}
              </div>

              {/* Active Evaluation Panel */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                
                {/* Writing Canvas */}
                <div className="border border-black p-5 bg-white bg-opacity-30 flex flex-col">
                  <div className="border-b border-black pb-2 mb-4">
                    <span className="text-[10px] font-mono uppercase tracking-wider font-bold text-red-950">
                      [Active Question: {selectedQuestion.marks} Marks]
                    </span>
                    <h3 className="font-serif font-bold text-base italic mt-1 text-black">
                      {selectedQuestion.question}
                    </h3>
                  </div>

                  <div className="bg-amber-50 border-l-4 border-amber-600 p-3 mb-4 text-xs font-mono">
                    <span className="font-bold text-amber-900 block mb-0.5">Professor's Last-minute Hint:</span>
                    <p className="text-amber-800 leading-normal">{selectedQuestion.hint}</p>
                  </div>

                  <label className="block text-[10px] font-mono uppercase tracking-wider font-bold mb-1">
                    Your Answer Sheet:
                  </label>
                  <textarea
                    rows={12}
                    value={userAnswer}
                    onChange={(e) => setUserAnswer(e.target.value)}
                    placeholder="Provide your exam answer following the standard pattern:
1. Definition
2. Syntax
3. Types/Features
4. Advantages
5. Detailed Example..."
                    className="w-full p-4 border border-black font-mono text-xs bg-[#E4E3E0] text-[#141414] focus:outline-none focus:ring-1 focus:ring-black placeholder-black/50 leading-relaxed"
                  />

                  <div className="mt-4 flex items-center justify-between">
                    <button
                      onClick={() => {
                        setUserAnswer(`1. Definition: ACID stands for Atomicity, Consistency, Isolation, and Durability. These properties ensure database transactions run safely and reliably.

2. Syntax/Model:
BEGIN TRANSACTION;
  -- DML Actions
COMMIT;

3. Types & Details:
- Atomicity: Executes completely or rolls back entirely (all or nothing).
- Consistency: DB moves from one valid state to another, obeying constraints.
- Isolation: Concurrent sessions cannot read dirty or uncommitted changes.
- Durability: Committed transaction data remains saved even if server crashes.

4. Advantages:
- Avoids financial transaction anomalies.
- Protects DB integrity against server power cuts.

5. Practical Example:
BEGIN;
  UPDATE accounts SET balance = balance - 100 WHERE id = 1;
  UPDATE accounts SET balance = balance + 100 WHERE id = 2;
COMMIT;`);
                      }}
                      className="text-xs font-mono underline hover:text-black/80"
                    >
                      Draft Sample Answer
                    </button>

                    <button
                      onClick={handleEvaluateAnswer}
                      disabled={isEvaluating || !userAnswer.trim()}
                      className="px-5 py-2.5 bg-black text-[#E4E3E0] hover:bg-black/95 font-mono text-xs uppercase font-bold tracking-wider flex items-center space-x-2 disabled:opacity-50"
                    >
                      {isEvaluating ? (
                        <>
                          <RefreshCw className="h-3.5 w-3.5 animate-spin" />
                          <span>Evaluating...</span>
                        </>
                      ) : (
                        <>
                          <UserCheck className="h-4 w-4 text-emerald-400" />
                          <span>Evaluate with AI</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>

                {/* AI Evaluation Report Sheet */}
                <div className="border border-black p-5 bg-black text-[#E4E3E0] flex flex-col justify-between">
                  <div>
                    <div className="border-b border-white/20 pb-2 mb-4 flex justify-between items-center font-mono">
                      <span className="text-[10px] text-emerald-400 uppercase font-bold">
                        [EVALUATION REPORT SHEET]
                      </span>
                      <span className="text-[10px] text-white/50">PostgreSQL Academic Evaluator</span>
                    </div>

                    {!evaluationResult && !isEvaluating ? (
                      <div className="py-20 text-center italic text-xs text-white/40 font-mono">
                        No answer graded yet. Write your answer on the left and submit it to see your marks breakdown.
                      </div>
                    ) : isEvaluating ? (
                      <div className="py-20 flex flex-col items-center justify-center space-y-3 font-mono text-xs">
                        <RefreshCw className="h-8 w-8 animate-spin text-emerald-400" />
                        <span className="text-white/60">Analyzing keyword mapping, checking PL/SQL syntax blocks, and matching normalized components...</span>
                      </div>
                    ) : (
                      <div className="space-y-4">
                        
                        {/* Score Block */}
                        <div className="flex items-center justify-between bg-white/5 p-4 border border-white/10">
                          <div>
                            <span className="font-mono text-[10px] text-white/50 block">OBTAINED MARKS</span>
                            <span className="text-2xl font-serif italic font-bold">
                              {evaluationResult.score} / 5 Marks
                            </span>
                          </div>
                          <div className="text-right">
                            <span className="font-mono text-[10px] text-white/50 block">VERDICT</span>
                            <span className={`text-xs font-bold font-mono px-2 py-0.5 rounded-none ${
                              evaluationResult.score >= 4 ? "bg-emerald-500 text-black" : "bg-amber-500 text-black"
                            }`}>
                              {evaluationResult.verdict}
                            </span>
                          </div>
                        </div>

                        {/* Strengths and Gaps */}
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          <div className="border border-white/10 p-3 bg-white/5">
                            <span className="font-mono text-[10px] text-emerald-400 font-bold block mb-1.5 uppercase">
                              ✔ What You Did Well:
                            </span>
                            <ul className="space-y-1 text-xs">
                              {evaluationResult.strengths?.map((str, i) => (
                                <li key={i} className="flex items-start">
                                  <span className="mr-1.5 text-emerald-400">•</span>
                                  <span>{str}</span>
                                </li>
                              ))}
                            </ul>
                          </div>

                          <div className="border border-white/10 p-3 bg-white/5">
                            <span className="font-mono text-[10px] text-red-400 font-bold block mb-1.5 uppercase">
                              ✗ Critical Gaps (Lost Marks):
                            </span>
                            <ul className="space-y-1 text-xs">
                              {evaluationResult.gaps?.map((gap, i) => (
                                <li key={i} className="flex items-start">
                                  <span className="mr-1.5 text-red-400">•</span>
                                  <span>{gap}</span>
                                </li>
                              ))}
                            </ul>
                          </div>
                        </div>

                        {/* Model answer reveal */}
                        <div className="mt-4 pt-3 border-t border-white/10">
                          <span className="font-mono text-[10px] text-white/60 block mb-1.5 uppercase">
                            Ideal Full-Marks Answer Blueprint:
                          </span>
                          <div className="max-h-48 overflow-y-auto text-xs font-sans text-white/80 bg-black/60 p-3 border border-white/10 leading-relaxed whitespace-pre-wrap">
                            {evaluationResult.modelAnswer}
                          </div>
                        </div>

                      </div>
                    )}
                  </div>

                  <div className="border-t border-white/10 mt-6 pt-3 text-[10px] text-white/40 text-right font-mono">
                    Evaluation based on university schema standards.
                  </div>
                </div>

              </div>

            </div>
          )}

          {/* TAB 4: VIVA & LAST-MINUTE TRICKS */}
          {activeTab === "viva" && (
            <div className="space-y-6">
              
              <div className="bg-black text-white p-5 tech-border">
                <span className="text-[10px] font-mono tracking-widest uppercase text-[#E4E3E0]/75">
                  Last-Minute Exam Hacks
                </span>
                <h2 className="text-xl font-serif font-bold italic mt-1 text-[#E4E3E0]">
                  Viva-Voce & Memory Mnemonics
                </h2>
                <p className="text-xs text-[#E4E3E0]/80 mt-1.5 leading-relaxed">
                  Fast track your memory using proven abbreviation formulas. Study typical tricks asked by external examiners during practical evaluations.
                </p>
              </div>

              {/* Memory Trick Cards - Bento Style */}
              <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
                {MEMORY_TRICKS.map((trick) => (
                  <div key={trick.id} className="border border-black bg-white bg-opacity-20 p-4 flex flex-col justify-between">
                    <div>
                      <span className="text-[10px] font-mono text-black/50 block mb-0.5">CONCEPT</span>
                      <h4 className="font-serif font-bold italic text-sm">{trick.concept}</h4>
                      <div className="text-lg font-mono font-black tracking-tight text-[#141414] my-2 bg-[#E4E3E0] p-1.5 border border-black/20 text-center uppercase">
                        {trick.acronym}
                      </div>
                      <p className="text-[11px] font-mono leading-tight mb-2 text-indigo-950 font-bold">{trick.expansion}</p>
                    </div>
                    <p className="text-xs opacity-75 leading-relaxed pt-2 border-t border-black/10">
                      {trick.explanation}
                    </p>
                  </div>
                ))}
              </div>

              {/* Tricky Viva Questions */}
              <div className="space-y-4 mt-6">
                <div className="border-b border-black pb-1.5 mb-4">
                  <span className="font-mono text-[10px] uppercase tracking-wider font-bold block">
                    TOP 5 VIVA VOCE QUESTIONS EXPECTED TOMORROW
                  </span>
                </div>

                <div className="divide-y divide-black/20 border border-black bg-[#E4E3E0]">
                  {VIVA_QUESTIONS.map((viva, index) => {
                    const isRevealed = revealedVivaIds.has(viva.id);
                    return (
                      <div key={viva.id} className="p-5 flex flex-col md:flex-row items-start justify-between bg-white bg-opacity-10">
                        <div className="flex-1 pr-6">
                          <span className="font-mono text-[10px] text-black/50 uppercase block">QUESTION {index + 1}</span>
                          <h4 className="font-serif font-bold text-base text-black mt-0.5 leading-tight">
                            {viva.question}
                          </h4>

                          {isRevealed && (
                            <div className="mt-3 space-y-2.5 font-sans">
                              <div className="text-xs text-black/85 leading-relaxed bg-black/5 p-3 border border-black/15">
                                <strong className="font-mono block text-[10px] text-black/60 mb-0.5">EXAMINER EXPECTED ANSWER:</strong>
                                {viva.answer}
                              </div>
                              <div className="text-[11px] text-emerald-800 bg-emerald-50 border-l-4 border-emerald-600 p-2 font-mono">
                                <strong>Examiner Pro Tip:</strong> {viva.proTip}
                              </div>
                            </div>
                          )}
                        </div>

                        <button
                          onClick={() => toggleVivaReveal(viva.id)}
                          className="mt-3 md:mt-0 font-mono text-xs px-3 py-1.5 border border-black uppercase font-bold hover:bg-black hover:text-[#E4E3E0] transition shrink-0"
                        >
                          {isRevealed ? "Hide Answer" : "Reveal Answer"}
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>

            </div>
          )}

          {/* TAB 5: NOTSOPEDIA COMMUNITY SHARING HUB */}
          {activeTab === "notesopedia" && (
            <div className="space-y-6">
              
              {/* Header Banner */}
              <div className="bg-black text-white p-5 tech-border flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div className="space-y-1.5 max-w-2xl">
                  <span className="text-[10px] font-mono tracking-widest uppercase text-emerald-400 font-bold flex items-center gap-1">
                    <Sparkles className="h-3 w-3" /> Notsopedia Peer-to-Peer Knowledge Hub
                  </span>
                  <h2 className="text-xl font-serif font-bold italic text-[#E4E3E0]">
                    Shared Course Repository
                  </h2>
                  <p className="text-xs text-[#E4E3E0]/80 leading-relaxed">
                    Access high-yield revision sheets, laboratory SQL notes, and lecture reviews. Feel free to upload your own database notes, exam templates, and helpful guides to assist others!
                  </p>
                </div>
                
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={() => {
                      if (!systemConfig.enableSubmissions && !isAdminMode) {
                        alert("⚠️ Student note submissions are currently locked by the System Administrator.");
                        return;
                      }
                      setShowUploadForm(!showUploadForm);
                      setNotesError("");
                      setNotesSuccessMessage("");
                    }}
                    className={`px-4 py-2 font-bold uppercase text-xs tracking-wider transition-all flex items-center justify-center space-x-1.5 shadow-[2px_2px_0px_rgba(0,0,0,1)] hover:shadow-[4px_4px_0px_rgba(0,0,0,1)] active:translate-x-[2px] active:translate-y-[2px] ${
                      !systemConfig.enableSubmissions && !isAdminMode 
                        ? "bg-neutral-500 text-neutral-800 cursor-not-allowed border-2 border-black" 
                        : "bg-[#10b981] hover:bg-emerald-600 text-black"
                    }`}
                  >
                    {showUploadForm ? (
                      <span>📚 View Notes Directory</span>
                    ) : (
                      <>
                        <UploadCloud className="h-4 w-4" />
                        <span>📤 Share Your Notes {!systemConfig.enableSubmissions && "(Locked)"}</span>
                      </>
                    )}
                  </button>

                  <button
                    onClick={() => {
                      if (isAdminMode) {
                        setIsAdminMode(false);
                        setNotesSuccessMessage("Admin mode deactivated.");
                        setTimeout(() => setNotesSuccessMessage(""), 3000);
                      } else {
                        setShowAdminPasswordModal(true);
                      }
                    }}
                    className={`px-4 py-2 border border-black font-mono font-bold uppercase text-xs tracking-wider transition-all flex items-center justify-center space-x-1.5 shadow-[2px_2px_0px_rgba(0,0,0,1)] hover:shadow-[4px_4px_0px_rgba(0,0,0,1)] active:translate-x-[2px] active:translate-y-[2px] ${
                      isAdminMode ? "bg-amber-400 text-black" : "bg-black text-white"
                    }`}
                  >
                    <Shield className="h-4 w-4" />
                    <span>{isAdminMode ? "🔒 Exit Admin" : "🔑 Admin Control"}</span>
                  </button>
                </div>
              </div>

              {/* Status / Success Messages */}
              {notesSuccessMessage && (
                <div className="p-4 bg-emerald-50 border-l-4 border-[#10b981] text-emerald-950 font-mono text-xs flex justify-between items-center rounded">
                  <div className="flex items-center space-x-2">
                    <span className="w-2 h-2 rounded-full bg-[#10b981] animate-ping"></span>
                    <span>{notesSuccessMessage}</span>
                  </div>
                  <button onClick={() => setNotesSuccessMessage("")} className="text-black font-bold hover:bg-black/10 px-1 rounded">✕</button>
                </div>
              )}

              {notesError && (
                <div className="p-4 bg-rose-50 border-l-4 border-rose-600 text-rose-950 font-mono text-xs flex justify-between items-center rounded">
                  <span>⚠️ {notesError}</span>
                  <button onClick={() => setNotesError("")} className="text-black font-bold hover:bg-black/10 px-1 rounded">✕</button>
                </div>
              )}

              {/* ADMIN CONTROL CONSOLE (Visible only when isAdminMode is true) */}
              {isAdminMode && (
                <div className="bg-amber-50 border-2 border-amber-500 p-5 shadow-[4px_4px_0px_rgba(0,0,0,1)] space-y-4 animate-fadeIn">
                  <div className="flex items-center justify-between border-b border-amber-500 pb-2">
                    <div className="flex items-center space-x-2">
                      <Shield className="h-5 w-5 text-amber-600 animate-pulse" />
                      <h3 className="text-sm font-mono font-black tracking-wider uppercase text-amber-900">
                        🛡️ Live App Administration & Control Console
                      </h3>
                    </div>
                    <span className="bg-amber-200 text-amber-900 text-[10px] font-mono px-2 py-0.5 rounded font-extrabold uppercase">
                      Admin Session Active
                    </span>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    {/* Control Card 1: Broadcast controls */}
                    <div className="bg-white p-4 border border-amber-300 space-y-3">
                      <span className="text-[10px] font-mono font-bold text-amber-950 block uppercase">
                        📢 Announcement Broadcaster
                      </span>
                      <textarea
                        rows={2}
                        value={systemConfig.announcement}
                        onChange={(e) => saveSystemConfig({ announcement: e.target.value })}
                        className="w-full p-2 border border-black font-mono text-xs focus:ring-1 focus:ring-amber-500 bg-neutral-50"
                        placeholder="Type a global custom announcement message..."
                      />
                      <label className="flex items-center space-x-2 text-xs font-mono select-none">
                        <input
                          type="checkbox"
                          checked={systemConfig.announcementActive}
                          onChange={(e) => saveSystemConfig({ announcementActive: e.target.checked })}
                          className="rounded text-amber-500 focus:ring-amber-500"
                        />
                        <span>Active Broadcast</span>
                      </label>
                    </div>

                    {/* Control Card 2: App Feature controls */}
                    <div className="bg-white p-4 border border-amber-300 space-y-3 flex flex-col justify-between">
                      <div>
                        <span className="text-[10px] font-mono font-bold text-amber-950 block uppercase">
                          ⚙️ App Feature Lock / Toggles
                        </span>
                        <div className="space-y-3 mt-2">
                          <label className="flex items-center justify-between text-xs font-mono select-none p-1.5 bg-neutral-50 border border-neutral-200">
                            <span>Enable Student Submissions</span>
                            <input
                              type="checkbox"
                              checked={systemConfig.enableSubmissions}
                              onChange={(e) => saveSystemConfig({ enableSubmissions: e.target.checked })}
                              className="rounded text-amber-500 focus:ring-amber-500 h-4 w-4"
                            />
                          </label>

                          <label className="flex items-center justify-between text-xs font-mono select-none p-1.5 bg-neutral-50 border border-neutral-200">
                            <span>Enable SQL Simulator</span>
                            <input
                              type="checkbox"
                              checked={systemConfig.enableSimulator}
                              onChange={(e) => saveSystemConfig({ enableSimulator: e.target.checked })}
                              className="rounded text-amber-500 focus:ring-amber-500 h-4 w-4"
                            />
                          </label>
                        </div>
                      </div>
                      <p className="text-[10px] text-amber-800 font-mono mt-2">
                        💡 Turning these toggles on/off instantly updates the backend, controlling client access in real-time.
                      </p>
                    </div>

                    {/* Control Card 3: Database Stats & Reset */}
                    <div className="bg-white p-4 border border-amber-300 space-y-3 flex flex-col justify-between">
                      <div>
                        <span className="text-[10px] font-mono font-bold text-amber-950 block uppercase">
                          📊 Real-time Data Metrics
                        </span>
                        <div className="grid grid-cols-2 gap-2 mt-2 font-mono text-xs">
                          <div className="bg-neutral-50 p-2 border border-neutral-200 text-center">
                            <span className="text-black/60 text-[9px] uppercase block">Total Notes</span>
                            <span className="text-base font-black">{userNotes.length}</span>
                          </div>
                          <div className="bg-neutral-50 p-2 border border-neutral-200 text-center">
                            <span className="text-black/60 text-[9px] uppercase block">Total Likes</span>
                            <span className="text-base font-black text-rose-600">
                              {userNotes.reduce((acc, curr) => acc + (curr.likes || 0), 0)}
                            </span>
                          </div>
                        </div>
                      </div>

                      <div className="pt-2 border-t border-amber-200 flex gap-2">
                        <button
                          type="button"
                          onClick={handleResetDatabase}
                          className="flex-1 px-2.5 py-1.5 bg-rose-600 hover:bg-rose-700 text-white font-bold font-mono text-[10px] uppercase text-center border border-black shadow-[1px_1px_0px_rgba(0,0,0,1)] transition-all"
                        >
                          ⚠️ Factory Reset
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setIsAdminMode(false);
                            setNotesSuccessMessage("Logged out of admin console.");
                            setTimeout(() => setNotesSuccessMessage(""), 3000);
                          }}
                          className="flex-1 px-2.5 py-1.5 bg-black text-[#E4E3E0] font-bold font-mono text-[10px] uppercase text-center border border-black shadow-[1px_1px_0px_rgba(0,0,0,1)] transition-all"
                        >
                          Lock Session
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* AUTHORIZE ADMIN MODAL */}
              {showAdminPasswordModal && (
                <div className="fixed inset-0 bg-black/75 backdrop-blur-xs flex items-center justify-center z-50 p-4">
                  <div className="bg-white border-4 border-black p-6 max-w-sm w-full shadow-[8px_8px_0px_rgba(0,0,0,1)] space-y-4 animate-fadeIn">
                    <div className="flex items-center space-x-2 text-black font-serif italic font-bold text-lg border-b-2 border-black pb-2">
                      <Shield className="h-5 w-5 text-amber-500" />
                      <span>Authorize Admin Session</span>
                    </div>
                    <p className="text-xs text-black/75 leading-relaxed font-mono">
                      Enter the administrative passcode to unlock database moderation, simulator toggles, and broadcast banner controls.
                      <br/>
                      <span className="text-emerald-700 font-bold block mt-1.5">🔑 Password is: admin</span>
                    </p>
                    <input
                      type="password"
                      placeholder="Enter passcode"
                      value={adminPassword}
                      onChange={(e) => setAdminPassword(e.target.value)}
                      className="w-full p-2.5 border-2 border-black font-mono text-sm bg-neutral-50 focus:bg-white focus:outline-none"
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          if (adminPassword === "admin") {
                            setIsAdminMode(true);
                            setShowAdminPasswordModal(false);
                            setAdminPassword("");
                            setNotesSuccessMessage("Welcome back, Administrator! Controls unlocked.");
                            setTimeout(() => setNotesSuccessMessage(""), 4000);
                          } else {
                            alert("Invalid administrative passcode.");
                          }
                        }
                      }}
                      autoFocus
                    />
                    <div className="flex justify-end space-x-2 font-mono text-xs pt-2">
                      <button
                        onClick={() => { setShowAdminPasswordModal(false); setAdminPassword(""); }}
                        className="px-3 py-1.5 border border-black hover:bg-neutral-100 uppercase"
                      >
                        Cancel
                      </button>
                      <button
                        onClick={() => {
                          if (adminPassword === "admin") {
                            setIsAdminMode(true);
                            setShowAdminPasswordModal(false);
                            setAdminPassword("");
                            setNotesSuccessMessage("Welcome back, Administrator! Controls unlocked.");
                            setTimeout(() => setNotesSuccessMessage(""), 4000);
                          } else {
                            alert("Invalid administrative passcode.");
                          }
                        }}
                        className="px-4 py-1.5 bg-black text-[#E4E3E0] hover:bg-neutral-900 font-bold uppercase"
                      >
                        Unlock 🔓
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* UPLOAD FORM */}
              {showUploadForm ? (
                <form onSubmit={handleAddNote} className="bg-white p-6 border-2 border-black shadow-[4px_4px_0px_0px_rgba(0,0,0,1)] space-y-5">
                  <div className="border-b border-black pb-2 mb-4">
                    <h3 className="text-base font-serif font-bold italic flex items-center gap-1.5">
                      <UploadCloud className="h-5 w-5 text-emerald-600" /> Share New Database Note
                    </h3>
                    <p className="text-[10px] font-mono opacity-60">Your submission will be immediately visible to anyone studying the curriculum.</p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {/* Note Title */}
                    <div className="space-y-1.5">
                      <label className="block text-xs font-mono font-bold uppercase">Note Title *</label>
                      <input
                        type="text"
                        placeholder="e.g., Transaction Isolation Levels Compared"
                        value={newNoteTitle}
                        onChange={(e) => setNewNoteTitle(e.target.value)}
                        className="w-full p-2.5 border border-black font-sans text-xs focus:ring-1 focus:ring-emerald-500 bg-neutral-50"
                        required
                      />
                    </div>

                    {/* Topic Name */}
                    <div className="space-y-1.5">
                      <label className="block text-xs font-mono font-bold uppercase">Topic / Unit Reference</label>
                      <input
                        type="text"
                        placeholder="e.g., Unit 4: Recovery, BCNF, etc."
                        value={newNoteTopicName}
                        onChange={(e) => setNewNoteTopicName(e.target.value)}
                        className="w-full p-2.5 border border-black font-sans text-xs focus:ring-1 focus:ring-emerald-500 bg-neutral-50"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {/* Subject Name */}
                    <div className="space-y-1.5">
                      <label className="block text-xs font-mono font-bold uppercase">Subject / Course Name *</label>
                      <select
                        value={newNoteSubjectName}
                        onChange={(e) => {
                          setNewNoteSubjectName(e.target.value);
                          if (e.target.value === "Relational Database Management Systems") {
                            setNewNoteSubjectCode("2505MJCT201");
                          } else if (e.target.value === "Database Systems & Administration") {
                            setNewNoteSubjectCode("CS-302");
                          } else {
                            setNewNoteSubjectCode("GEN-DB");
                          }
                        }}
                        className="w-full p-2.5 border border-black font-sans text-xs bg-white"
                      >
                        <option value="Relational Database Management Systems">Relational Database Management Systems</option>
                        <option value="Database Systems & Administration">Database Systems & Administration</option>
                        <option value="Other Database Elective">Other Database Elective / Custom Course</option>
                      </select>
                    </div>

                    {/* Subject Code */}
                    <div className="space-y-1.5">
                      <label className="block text-xs font-mono font-bold uppercase">Subject Code</label>
                      <input
                        type="text"
                        placeholder="e.g., 2505MJCT201"
                        value={newNoteSubjectCode}
                        onChange={(e) => setNewNoteSubjectCode(e.target.value)}
                        className="w-full p-2.5 border border-black font-sans text-xs focus:ring-1 focus:ring-emerald-500 bg-neutral-50"
                      />
                    </div>
                  </div>

                  {/* Uploader Details Section */}
                  <div className="p-4 bg-emerald-50/50 border border-emerald-900/10 space-y-4">
                    <span className="text-[10px] font-mono tracking-wider font-bold block uppercase text-emerald-900/80">
                      👤 UPLOADER PROFILE INFORMATION
                    </span>
                    
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                      {/* Uploader Name */}
                      <div className="space-y-1.5">
                        <label className="block text-[11px] font-mono font-bold uppercase">Your Name *</label>
                        <input
                          type="text"
                          placeholder="e.g., Chetana Agle"
                          value={newNoteUploaderName}
                          onChange={(e) => setNewNoteUploaderName(e.target.value)}
                          className="w-full p-2 border border-black font-sans text-xs bg-white"
                          required
                        />
                      </div>

                      {/* Uploader Role */}
                      <div className="space-y-1.5">
                        <label className="block text-[11px] font-mono font-bold uppercase">Academic Role *</label>
                        <select
                          value={newNoteUploaderRole}
                          onChange={(e) => setNewNoteUploaderRole(e.target.value)}
                          className="w-full p-2 border border-black font-sans text-xs bg-white"
                        >
                          <option value="Student">Student</option>
                          <option value="Lead TA">Lead TA / Tutor</option>
                          <option value="Teaching Assistant">Teaching Assistant</option>
                          <option value="Professor">Professor / Instructor</option>
                          <option value="Alumni">Alumni</option>
                        </select>
                      </div>

                      {/* Uploader Email */}
                      <div className="space-y-1.5">
                        <label className="block text-[11px] font-mono font-bold uppercase">Email (Optional)</label>
                        <input
                          type="email"
                          placeholder="e.g., user@university.edu"
                          value={newNoteUploaderEmail}
                          onChange={(e) => setNewNoteUploaderEmail(e.target.value)}
                          className="w-full p-2 border border-black font-sans text-xs bg-white"
                        />
                      </div>
                    </div>
                  </div>

                  {/* Notes Content */}
                  <div className="space-y-1.5">
                    <label className="block text-xs font-mono font-bold uppercase">Notes Content * (Supports Multi-line / Markdown formatting)</label>
                    <textarea
                      rows={10}
                      placeholder="Write your comprehensive revision notes here. You can use markdown like ### Subheadings or - Bullet lists..."
                      value={newNoteContent}
                      onChange={(e) => setNewNoteContent(e.target.value)}
                      className="w-full p-3 border border-black font-mono text-xs focus:ring-1 focus:ring-emerald-500 bg-neutral-50"
                      required
                    ></textarea>
                    <p className="text-[10px] text-black/50 font-mono">⚠️ Notes containing offensive or incorrect statements are subject to moderator flagging.</p>
                  </div>

                  {/* Submit buttons */}
                  <div className="flex justify-end space-x-2 pt-2 border-t border-black/10">
                    <button
                      type="button"
                      onClick={() => setShowUploadForm(false)}
                      className="px-4 py-2 border border-black font-mono text-xs uppercase hover:bg-neutral-100"
                    >
                      Cancel
                    </button>
                    <button
                      type="submit"
                      disabled={isSubmittingNote}
                      className="px-6 py-2 bg-black text-[#E4E3E0] font-bold font-mono text-xs uppercase hover:bg-neutral-900 transition flex items-center space-x-2 disabled:opacity-50"
                    >
                      {isSubmittingNote ? <span>Publishing...</span> : <span>🚀 Publish Note</span>}
                    </button>
                  </div>
                </form>
              ) : (
                <>
                  {/* DIRECTORY VIEW */}
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-4 border border-black bg-white bg-opacity-30">
                    {/* Search */}
                    <div className="flex-1 relative">
                      <Search className="absolute left-3 top-3 h-4 w-4 opacity-40" />
                      <input
                        type="text"
                        placeholder="Search notes by subject, title, uploader name, or content..."
                        value={notesSearchQuery}
                        onChange={(e) => setNotesSearchQuery(e.target.value)}
                        className="w-full pl-10 pr-4 py-2 border border-black bg-white text-xs font-mono"
                      />
                    </div>
                    
                    {/* Subject Filter */}
                    <div className="flex items-center space-x-2">
                      <label className="text-xs font-mono font-bold shrink-0">Filter Course:</label>
                      <select
                        value={selectedSubjectFilter}
                        onChange={(e) => setSelectedSubjectFilter(e.target.value)}
                        className="p-2 border border-black text-xs font-mono bg-white"
                      >
                        <option value="All Subjects">All Subjects</option>
                        <option value="Relational Database Management Systems">Relational Database Management Systems</option>
                        <option value="Database Systems & Administration">Database Systems & Administration</option>
                      </select>
                    </div>

                    <button 
                      onClick={() => fetchNotes()}
                      disabled={isLoadingNotes}
                      className="p-2 border border-black text-xs hover:bg-black/5 disabled:opacity-50 shrink-0 flex items-center justify-center gap-1.5 font-mono"
                      title="Sync from server"
                    >
                      <RefreshCw className={`h-3 w-3 ${isLoadingNotes ? 'animate-spin' : ''}`} />
                      <span>Refresh</span>
                    </button>
                  </div>

                  {/* Notes Grid */}
                  {isLoadingNotes ? (
                    <div className="py-12 text-center font-mono text-xs space-y-2">
                      <RefreshCw className="h-6 w-6 animate-spin mx-auto text-emerald-600" />
                      <span>Syncing Notesopedia archives...</span>
                    </div>
                  ) : (
                    (() => {
                      // Filter and search logic
                      const filteredNotes = userNotes.filter(note => {
                        const matchesFilter = selectedSubjectFilter === "All Subjects" || note.subjectName === selectedSubjectFilter;
                        
                        const query = notesSearchQuery.toLowerCase();
                        const matchesSearch = !notesSearchQuery ||
                          note.title.toLowerCase().includes(query) ||
                          note.content.toLowerCase().includes(query) ||
                          note.subjectName.toLowerCase().includes(query) ||
                          (note.subjectCode && note.subjectCode.toLowerCase().includes(query)) ||
                          note.uploaderName.toLowerCase().includes(query) ||
                          (note.topicName && note.topicName.toLowerCase().includes(query));
                        
                        return matchesFilter && matchesSearch;
                      });

                      if (filteredNotes.length === 0) {
                        return (
                          <div className="border border-black border-dashed p-10 text-center text-xs font-mono space-y-3">
                            <div>No notes match your search or filter.</div>
                            <button
                              onClick={() => { setNotesSearchQuery(""); setSelectedSubjectFilter("All Subjects"); }}
                              className="px-3 py-1 bg-black text-white uppercase text-[10px]"
                            >
                              Clear Filters
                            </button>
                          </div>
                        );
                      }

                      return (
                        <div className="grid grid-cols-1 gap-6">
                          {filteredNotes.map((note) => {
                            const isExpanded = expandedNoteId === note.id;
                            const previewText = note.content.split("\n").slice(0, 4).join("\n");
                            
                            return (
                              <div key={note.id} className="bg-white border-2 border-black shadow-[4px_4px_0px_rgba(0,0,0,1)] flex flex-col hover:shadow-[6px_6px_0px_rgba(0,0,0,1)] transition-all">
                                
                                {/* Card Header */}
                                <div className="border-b border-black p-4 bg-emerald-50/20 flex flex-col md:flex-row md:items-center justify-between gap-3">
                                  <div className="space-y-1">
                                    <div className="flex flex-wrap items-center gap-1.5 font-mono text-[10px]">
                                      <span className="px-2 py-0.5 bg-black text-white rounded font-bold uppercase tracking-wider">
                                        {note.subjectCode}
                                      </span>
                                      <span className="text-black/60 font-semibold uppercase flex items-center gap-1">
                                        <Folder className="h-3 w-3 inline" /> {note.subjectName}
                                      </span>
                                      {note.topicName && (
                                        <span className="text-indigo-900 font-bold bg-indigo-50 px-1.5 py-0.2 rounded border border-indigo-900/10 flex items-center gap-1">
                                          <Tag className="h-3 w-3 inline text-indigo-700" /> {note.topicName}
                                        </span>
                                      )}
                                    </div>
                                    <h3 className="text-lg font-serif font-black italic text-black leading-tight">
                                      {note.title}
                                    </h3>
                                  </div>

                                  {/* Uploader Profile Badge */}
                                  <div className="bg-[#E4E3E0] border border-black p-2.5 font-mono text-[11px] leading-tight flex items-center space-x-3 shrink-0">
                                    <div className="bg-black text-white rounded-full w-7 h-7 flex items-center justify-center font-bold">
                                      {note.uploaderName.charAt(0)}
                                    </div>
                                    <div>
                                      <div className="font-bold flex items-center gap-1 text-neutral-900">
                                        <User className="h-3.5 w-3.5 inline text-neutral-700" /> {note.uploaderName}
                                      </div>
                                      <div className="text-[10px] opacity-75 flex items-center justify-between gap-1.5 mt-0.5">
                                        <span className="bg-emerald-100 text-emerald-950 px-1 py-0.2 text-[9px] rounded font-bold uppercase">{note.uploaderRole}</span>
                                        <span>• {new Date(note.uploadedAt).toLocaleDateString()}</span>
                                      </div>
                                      {note.uploaderEmail && (
                                        <div className="text-[10px] text-neutral-600 flex items-center gap-1 mt-0.5">
                                          <Mail className="h-3.5 w-3.5 inline text-neutral-500" /> {note.uploaderEmail}
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                </div>

                                {/* Card Body (Notes Content) */}
                                <div className="p-5 font-sans text-xs leading-relaxed text-black/90 whitespace-pre-wrap flex-1 bg-white">
                                  {isExpanded ? (
                                    <div className="space-y-4">
                                      {note.content}
                                    </div>
                                  ) : (
                                    <div>
                                      {previewText}...
                                      {note.content.length > previewText.length && (
                                        <button
                                          type="button"
                                          onClick={() => setExpandedNoteId(note.id)}
                                          className="text-emerald-700 font-bold ml-1.5 hover:underline"
                                        >
                                          [Read Full Note]
                                        </button>
                                      )}
                                    </div>
                                  )}
                                </div>

                                {/* Card Footer Controls */}
                                <div className="border-t border-black px-5 py-3.5 bg-neutral-50 flex items-center justify-between flex-wrap gap-3 font-mono text-xs">
                                  <div className="flex items-center space-x-3">
                                    <button
                                      type="button"
                                      onClick={() => handleLikeNote(note.id)}
                                      className="flex items-center space-x-1.5 px-3 py-1 bg-rose-50 text-rose-950 border border-black hover:bg-rose-100 active:translate-y-[0.5px] transition"
                                      title="Show appreciation for this note!"
                                    >
                                      <Heart className="h-3.5 w-3.5 text-rose-600 fill-rose-600" />
                                      <span className="font-bold">Helpful? ({note.likes || 0})</span>
                                    </button>
                                    
                                    <button
                                      type="button"
                                      onClick={() => {
                                        navigator.clipboard.writeText(`${note.title}\n\nSubject: ${note.subjectName}\nTopic: ${note.topicName}\nShared by: ${note.uploaderName} (${note.uploaderRole})\n\n${note.content}`);
                                        alert("Full note copied to clipboard!");
                                      }}
                                      className="px-2.5 py-1 border border-black bg-white hover:bg-neutral-100 transition"
                                    >
                                      📋 Copy Text
                                    </button>

                                    {isAdminMode && (
                                      <div className="flex items-center space-x-2 border-l border-black/20 pl-3">
                                        <button
                                          type="button"
                                          onClick={() => setEditingNote(note)}
                                          className="flex items-center space-x-1 px-2 py-1 bg-amber-100 text-amber-950 border border-black hover:bg-amber-200 transition"
                                          title="Edit this note"
                                        >
                                          <Edit3 className="h-3 w-3 text-amber-800" />
                                          <span className="font-bold text-[10px]">Edit</span>
                                        </button>
                                        <button
                                          type="button"
                                          onClick={() => handleDeleteNote(note.id)}
                                          className="flex items-center space-x-1 px-2 py-1 bg-rose-600 text-white border border-black hover:bg-rose-700 transition"
                                          title="Delete this note permanently"
                                        >
                                          <Trash2 className="h-3 w-3 text-white" />
                                          <span className="font-bold text-[10px]">Delete</span>
                                        </button>
                                      </div>
                                    )}
                                  </div>

                                  {isExpanded && (
                                    <button
                                      type="button"
                                      onClick={() => setExpandedNoteId(null)}
                                      className="text-black hover:underline font-bold text-[11px]"
                                    >
                                      Collapse Note ▴
                                    </button>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      );
                    })()
                  )}
                </>
              )}

              {/* EDITING NOTE OVERLAY MODAL */}
              {editingNote && (
                <div className="fixed inset-0 bg-black/70 backdrop-blur-xs flex items-center justify-center z-50 p-4">
                  <form onSubmit={handleEditNoteSubmit} className="bg-white border-4 border-black p-6 max-w-lg w-full shadow-[8px_8px_0px_rgba(0,0,0,1)] space-y-4 animate-fadeIn">
                    <div className="flex items-center justify-between border-b-2 border-black pb-2">
                      <div className="flex items-center space-x-2 text-black font-serif italic font-bold text-lg">
                        <Edit3 className="h-5 w-5 text-emerald-600 animate-pulse" />
                        <span>Edit Shared Peer Note</span>
                      </div>
                      <span className="text-[10px] font-mono bg-black text-[#E4E3E0] px-2 py-0.5 rounded uppercase font-bold">
                        ID: {editingNote.id.slice(0, 8)}
                      </span>
                    </div>

                    <div className="space-y-3 font-mono text-xs text-black">
                      {/* Note Title */}
                      <div className="space-y-1">
                        <label className="block font-bold uppercase">Note Title *</label>
                        <input
                          type="text"
                          value={editingNote.title}
                          onChange={(e) => setEditingNote(prev => prev ? { ...prev, title: e.target.value } : null)}
                          className="w-full p-2.5 border border-black font-sans text-xs bg-neutral-50"
                          required
                        />
                      </div>

                      <div className="grid grid-cols-2 gap-2">
                        {/* Subject Name */}
                        <div className="space-y-1">
                          <label className="block font-bold uppercase">Subject Name *</label>
                          <select
                            value={editingNote.subjectName}
                            onChange={(e) => setEditingNote(prev => prev ? { ...prev, subjectName: e.target.value, subjectCode: e.target.value.includes("Administration") ? "2505MJCT202" : "2505MJCT201" } : null)}
                            className="w-full p-2.5 border border-black bg-neutral-50 text-[11px]"
                            required
                          >
                            <option value="Relational Database Management Systems">Relational Database Management Systems (2505MJCT201)</option>
                            <option value="Database Systems & Administration">Database Systems & Administration (2505MJCT202)</option>
                          </select>
                        </div>

                        {/* Topic Reference */}
                        <div className="space-y-1">
                          <label className="block font-bold uppercase">Topic/Unit Reference</label>
                          <input
                            type="text"
                            value={editingNote.topicName || ""}
                            onChange={(e) => setEditingNote(prev => prev ? { ...prev, topicName: e.target.value } : null)}
                            className="w-full p-2.5 border border-black bg-neutral-50"
                            placeholder="e.g. Unit 2 - PL/SQL triggers"
                          />
                        </div>
                      </div>

                      {/* Note Content */}
                      <div className="space-y-1">
                        <label className="block font-bold uppercase">Note Content (Markdown Supported) *</label>
                        <textarea
                          rows={6}
                          value={editingNote.content}
                          onChange={(e) => setEditingNote(prev => prev ? { ...prev, content: e.target.value } : null)}
                          className="w-full p-2.5 border border-black font-sans text-xs bg-neutral-50 leading-relaxed"
                          required
                        />
                      </div>

                      <div className="grid grid-cols-2 gap-2">
                        {/* Uploader Name */}
                        <div className="space-y-1">
                          <label className="block font-bold uppercase">Uploader Name *</label>
                          <input
                            type="text"
                            value={editingNote.uploaderName}
                            onChange={(e) => setEditingNote(prev => prev ? { ...prev, uploaderName: e.target.value } : null)}
                            className="w-full p-2.5 border border-black bg-neutral-50"
                            required
                          />
                        </div>

                        {/* Uploader Role */}
                        <div className="space-y-1">
                          <label className="block font-bold uppercase">Role *</label>
                          <select
                            value={editingNote.uploaderRole}
                            onChange={(e) => setEditingNote(prev => prev ? { ...prev, uploaderRole: e.target.value } : null)}
                            className="w-full p-2.5 border border-black bg-neutral-50"
                          >
                            <option value="Student">Student</option>
                            <option value="Representative">Class Representative</option>
                            <option value="Professor">Faculty/Professor</option>
                          </select>
                        </div>
                      </div>
                    </div>

                    <div className="flex justify-end space-x-2 font-mono text-xs pt-2">
                      <button
                        type="button"
                        onClick={() => setEditingNote(null)}
                        className="px-4 py-2 border border-black hover:bg-neutral-100 uppercase"
                      >
                        Cancel
                      </button>
                      <button
                        type="submit"
                        className="px-5 py-2 bg-[#10b981] hover:bg-emerald-600 text-black font-bold uppercase"
                      >
                        Save Note Changes 💾
                      </button>
                    </div>
                  </form>
                </div>
              )}
            </div>
          )}

        </main>

        {/* RIGHT HAND QUICK PANEL */}
        <aside className="w-80 flex flex-col bg-white/40 border-l border-black h-full overflow-y-auto">
          
          {/* Quick Info Box */}
          <div className="p-4 border-b border-black bg-black text-[#E4E3E0]">
            <span className="text-[10px] font-mono text-emerald-400 uppercase font-bold block">
              [CRITICAL TOPIC WATCH]
            </span>
            <div className="font-serif font-bold italic text-base mt-1 text-white">
              Syllabus High Priority
            </div>
            <p className="text-xs text-[#E4E3E0]/70 mt-1.5 leading-relaxed">
              We anticipate 2PL, ACID, log-based recovery (deferred modification), and Triggers will encompass 60% of the entire tomorrow's exam marks.
            </p>
          </div>

          {/* Core Mnemonics Mini Block */}
          <div className="p-4 border-b border-black">
            <span className="text-[10px] font-mono uppercase tracking-wider font-bold block mb-2.5">
              Rapid Mnemonic Codes
            </span>
            <div className="space-y-3 font-mono text-xs">
              <div className="flex justify-between items-center border-b border-black/10 pb-1.5">
                <span className="font-bold">A.C.I.D.</span>
                <span className="text-[10px] opacity-60">Atom, Consist, Iso, Dur</span>
              </div>
              <div className="flex justify-between items-center border-b border-black/10 pb-1.5">
                <span className="font-bold">T.C.L.</span>
                <span className="text-[10px] opacity-60">Commit, Rollback, Save</span>
              </div>
              <div className="flex justify-between items-center border-b border-black/10 pb-1.5">
                <span className="font-bold">D.C.L.</span>
                <span className="text-[10px] opacity-60">Grant, Revoke</span>
              </div>
              <div className="flex justify-between items-center pb-1.5">
                <span className="font-bold">W-D vs W-W</span>
                <span className="text-[10px] opacity-60">Wait-Die vs Wound-Wait</span>
              </div>
            </div>
          </div>

          {/* Expected list */}
          <div className="p-4 border-b border-black">
            <span className="text-[10px] font-mono uppercase tracking-wider font-bold block mb-2">
              Expected 5-Mark List
            </span>
            <ul className="text-xs space-y-1.5 font-sans text-black/80">
              <li className="flex items-center space-x-1.5">
                <span className="w-1.5 h-1.5 bg-black rounded-full"></span>
                <span>Procedure vs Function</span>
              </li>
              <li className="flex items-center space-x-1.5">
                <span className="w-1.5 h-1.5 bg-black rounded-full"></span>
                <span>Explain Log-Based Recovery</span>
              </li>
              <li className="flex items-center space-x-1.5">
                <span className="w-1.5 h-1.5 bg-black rounded-full"></span>
                <span>Two-Phase Locking Protocol</span>
              </li>
              <li className="flex items-center space-x-1.5">
                <span className="w-1.5 h-1.5 bg-black rounded-full"></span>
                <span>Define Triggers & Write Syntax</span>
              </li>
              <li className="flex items-center space-x-1.5">
                <span className="w-1.5 h-1.5 bg-black rounded-full"></span>
                <span>Compare 1NF, 2NF, 3NF</span>
              </li>
            </ul>
          </div>

          {/* Running Engine logs mimicking system activity cleanly */}
          <div className="p-4 font-mono text-[10px] leading-tight space-y-1.5 mt-auto border-t border-black bg-black text-emerald-400">
            <div>SYSTEM LOGS:</div>
            <div>&gt; Initializing exam booster...</div>
            <div>&gt; Syllabus loaded: Units 1-6</div>
            <div>&gt; Connection status: OPTIMISTIC</div>
            <div className="text-white/50 text-right text-[8px] pt-1">
              v2.0 • University Edition
            </div>
          </div>

        </aside>

      </div>
    </div>
  );
}
