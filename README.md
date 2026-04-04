const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, BorderStyle, HeadingLevel, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const ACCENT = "1A56DB";   // Professional blue
const DARK   = "111827";
const GRAY   = "6B7280";

function sectionHeader(text) {
  return new Paragraph({
    spacing: { before: 180, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 4 } },
    children: [
      new TextRun({ text, bold: true, size: 22, color: ACCENT, font: "Arial" })
    ]
  });
}

function jobHeader(title, company, location, dates) {
  return new Paragraph({
    spacing: { before: 120, after: 20 },
    tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
    children: [
      new TextRun({ text: company, bold: true, size: 20, color: DARK, font: "Arial" }),
      new TextRun({ text: " | ", size: 20, color: GRAY, font: "Arial" }),
      new TextRun({ text: title, bold: true, size: 20, color: ACCENT, font: "Arial" }),
      new TextRun({ text: "\t", size: 20, font: "Arial" }),
      new TextRun({ text: `${location}  |  ${dates}`, size: 18, color: GRAY, italics: true, font: "Arial" }),
    ]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 20, after: 20 },
    children: [new TextRun({ text, size: 18, color: DARK, font: "Arial" })]
  });
}

function projHeader(name, stack) {
  return new Paragraph({
    spacing: { before: 120, after: 20 },
    children: [
      new TextRun({ text: name, bold: true, size: 20, color: DARK, font: "Arial" }),
      new TextRun({ text: "  |  ", size: 18, color: GRAY, font: "Arial" }),
      new TextRun({ text: stack, size: 18, color: GRAY, italics: true, font: "Arial" }),
    ]
  });
}

function skillRow(label, value) {
  return new Paragraph({
    spacing: { before: 20, after: 20 },
    children: [
      new TextRun({ text: label + ":  ", bold: true, size: 18, color: DARK, font: "Arial" }),
      new TextRun({ text: value, size: 18, color: DARK, font: "Arial" }),
    ]
  });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 260 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20, color: DARK } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 864, right: 1008, bottom: 864, left: 1008 }
      }
    },
    children: [
      // ── NAME ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: "VRAJESH RAKESHBHAI SHAH", bold: true, size: 36, color: DARK, font: "Arial" })
        ]
      }),
      // ── TAGLINE ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [
          new TextRun({ text: "Software Engineer  ·  Blockchain & Web3 Developer  ·  AI/ML Enthusiast", size: 18, color: GRAY, italics: true, font: "Arial" })
        ]
      }),
      // ── CONTACT ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 160 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "D1D5DB", space: 8 } },
        children: [
          new TextRun({ text: "+1 (213) 824-6088", size: 18, color: GRAY, font: "Arial" }),
          new TextRun({ text: "  ·  ", size: 18, color: GRAY, font: "Arial" }),
          new TextRun({ text: "vrajesh.web3@gmail.com", size: 18, color: GRAY, font: "Arial" }),
          new TextRun({ text: "  ·  ", size: 18, color: GRAY, font: "Arial" }),
          new TextRun({ text: "linkedin.com/in/vrsx", size: 18, color: GRAY, font: "Arial" }),
          new TextRun({ text: "  ·  ", size: 18, color: GRAY, font: "Arial" }),
          new TextRun({ text: "San Jose, CA 95134  ·  Open to Relocation  ·  Immediate Start", size: 18, color: GRAY, font: "Arial" }),
        ]
      }),

      // ── ABOUT ──
      sectionHeader("ABOUT ME"),
      new Paragraph({
        spacing: { before: 60, after: 80 },
        children: [new TextRun({
          text: "Passionate software engineer focused on blockchain development, decentralized applications (dApps), and smart contracts. Deeply interested in how decentralized technologies and DeFi are reshaping industries — from compliance infrastructure to real-time fraud detection. I bring full-stack and distributed systems expertise to every project, combining production-grade backend engineering with a drive to build impactful, trustless systems.",
          size: 18, color: DARK, font: "Arial"
        })]
      }),

      // ── SKILLS ──
      sectionHeader("TECHNICAL SKILLS"),
      skillRow("Languages", "C++, Python, Scala, Java, Go, SQL, TypeScript, JavaScript, Solidity"),
      skillRow("Blockchain & Web3", "Solidity, Web3.js, Smart Contracts, dApps, DeFi, Ethereum, Hardhat, IPFS"),
      skillRow("Systems & Backend", "Scala (gRPC), Spring Boot (Hibernate), Node.js, FastAPI, Kafka, GraphQL, RabbitMQ"),
      skillRow("AI & ML", "PyTorch, TensorFlow, RAG (Pinecone), Scikit-learn, SHAP/LIME"),
      skillRow("Cloud & DevOps", "AWS (GPU Clusters, Lambda, EC2, S3), Kubernetes, Docker, Prometheus, GitHub Actions, Jenkins"),
      skillRow("Frontend", "React.js, Next.js, React Native, Vue.js, TailwindCSS, Three.js, WebGL"),
      skillRow("Databases", "PostgreSQL, MongoDB, Redis, MySQL, DynamoDB, Apache Cassandra"),

      // ── EXPERIENCE ──
      sectionHeader("EXPERIENCE"),

      jobHeader("Software Engineer Intern — Compliance Engineering", "Gemini", "New York, NY", "May 2025 – Aug 2025"),
      bullet("Designed high-throughput Scala/gRPC services backed by PostgreSQL, integrating Hummingbird and Chainalysis; resolved a critical incident affecting 14K+ U.S./EU users, improving fraud detection accuracy by 28% and cutting audit time by 35%."),
      bullet("Re-architected HubAdmin into a distributed Refine + React frontend with a Scala gRPC backend and Node.js services, enabling scalable account intelligence workflows and reducing analyst overhead by 40%."),
      bullet("Built automated regression pipelines with Playwright, Jenkins, GitHub Actions, and AWS (EC2, S3, Lambda), strengthening release stability across compliance-critical flows and reducing regression latency by 45%."),

      jobHeader("Research Assistant — Applied AI & ML Research", "CalState Los Angeles", "Los Angeles, CA", "Oct 2024 – Apr 2025"),
      bullet("Contributed to fault-tolerant distributed training pipelines for Transformer and CNN models using PyTorch and TensorFlow, orchestrating multi-node AWS GPU workloads via Docker to cut latency by 25%."),
      bullet("Built high-throughput inference microservices with Python and FastAPI, optimizing data ingestion via PostgreSQL and NumPy to maintain 93% accuracy; integrated SHAP and LIME for model explainability in production AI environments."),

      jobHeader("Software Engineer", "DevKrutiTech", "Remote", "Apr 2023 – May 2024"),
      bullet("Migrated 12 core REST endpoints to a unified GraphQL schema backed by Spring Boot microservices, reducing client-side codebase by 25% and reliably handling 15K+ requests."),
      bullet("Architected fault-tolerant data pipelines using Kafka, RabbitMQ, and Hibernate to decouple backend bottlenecks — driving a 40% increase in application throughput under peak loads."),
      bullet("Engineered Redis and PostgreSQL caching layers, slashing system latency by ~300ms and reducing infrastructure scaling costs by 20%."),

      // ── PROJECTS ──
      sectionHeader("PROJECTS"),

      projHeader("OmniRAG: Context-Aware Enterprise LLM Agent", "Python, PyTorch, FastAPI, Pinecone, React.js, Docker"),
      bullet("Engineered a high-throughput RAG pipeline using PyTorch, FastAPI, and Pinecone vector databases, mitigating LLM hallucinations by 35% across domain-specific enterprise queries."),
      bullet("Architected a fault-tolerant Dockerized React.js frontend serving 10K+ concurrent requests with sub-second query latency — secured 1st place among 50+ competing engineering teams."),

      projHeader("LedgerGuard: Real-Time Financial Fraud Detection", "Java, Spring Boot, Kafka, PostgreSQL, Scikit-learn, AWS"),
      bullet("Architected a distributed financial anomaly detection engine using Java, Spring Boot, and Apache Kafka to asynchronously ingest and process 1M+ daily distributed ledger transactions."),
      bullet("Integrated a Scikit-learn ML model with PostgreSQL, achieving 96% fraud classification accuracy with 99.9% uptime across autoscaling AWS EC2 instances."),

      projHeader("KubeServe: Distributed AI Model Serving Platform", "Kubernetes, Go, Docker, Prometheus, AWS (EC2, GPU)"),
      bullet("Developed a cloud-native ML model serving platform in Go, leveraging Kubernetes HPA to dynamically provision AWS GPU nodes based on real-time inference traffic."),
      bullet("Implemented observability and CI/CD pipelines with Prometheus and Docker, enabling zero-downtime deployments and reducing compute costs by 30%."),

      // ── EDUCATION ──
      sectionHeader("EDUCATION"),
      new Paragraph({
        spacing: { before: 100, after: 20 },
        tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
        children: [
          new TextRun({ text: "California State University, Los Angeles", bold: true, size: 20, color: DARK, font: "Arial" }),
          new TextRun({ text: "\t", size: 20, font: "Arial" }),
          new TextRun({ text: "Aug 2024 – Jan 2026", size: 18, color: GRAY, italics: true, font: "Arial" }),
        ]
      }),
      new Paragraph({
        spacing: { before: 0, after: 20 },
        children: [new TextRun({ text: "Master of Science in Computer Science  ·  Los Angeles, CA", size: 18, color: GRAY, font: "Arial" })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 20 },
        children: [new TextRun({ text: "Coursework: Software Engineering, Artificial Intelligence, Distributed Systems, Data Science, Information Security, Database Management, Computer Graphics, Operating Systems, Software Architecture, Web Programming", size: 17, color: GRAY, italics: true, font: "Arial" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/mnt/user-data/outputs/Vrajesh_Shah_Resume.docx", buf);
  console.log("Done");
});
