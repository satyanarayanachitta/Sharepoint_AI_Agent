(() => {
  var __create = Object.create;
  var __defProp = Object.defineProperty;
  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __getProtoOf = Object.getPrototypeOf;
  var __hasOwnProp = Object.prototype.hasOwnProperty;
  var __require = /* @__PURE__ */ ((x) => typeof require !== "undefined" ? require : typeof Proxy !== "undefined" ? new Proxy(x, {
    get: (a, b) => (typeof require !== "undefined" ? require : a)[b]
  }) : x)(function(x) {
    if (typeof require !== "undefined") return require.apply(this, arguments);
    throw Error('Dynamic require of "' + x + '" is not supported');
  });
  var __copyProps = (to, from, except, desc) => {
    if (from && typeof from === "object" || typeof from === "function") {
      for (let key of __getOwnPropNames(from))
        if (!__hasOwnProp.call(to, key) && key !== except)
          __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
    }
    return to;
  };
  var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
    // If the importer is in node compatibility mode or this is not an ESM
    // file that has been converted to a CommonJS file using a Babel-
    // compatible transform (i.e. "__esModule" has not been set), then set
    // "default" to the CommonJS "module.exports" for node compatibility.
    isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
    mod
  ));

  // App.jsx
  var import_react = __toESM(__require("react"));
  var initialConfig = {
    siteUrl: "https://your-sharepoint-site-url/sites/your-site-name",
    folderPath: "Shared Documents/Your-Folder-Name",
    openAIApiKey: "",
    username: "",
    password: ""
  };
  function App() {
    const [config, setConfig] = (0, import_react.useState)(initialConfig);
    const [loading, setLoading] = (0, import_react.useState)(false);
    const [results, setResults] = (0, import_react.useState)([]);
    const [error, setError] = (0, import_react.useState)("");
    const handleChange = (event) => {
      const { name, value } = event.target;
      setConfig((prev) => ({ ...prev, [name]: value }));
    };
    const getDocumentsFromSharePoint = async () => {
      const response = await fetch("/api/sharepoint/documents", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          siteUrl: config.siteUrl,
          folderPath: config.folderPath,
          username: config.username,
          password: config.password
        })
      });
      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.message || "Unable to fetch documents from SharePoint.");
      }
      return response.json();
    };
    const analyzeDocumentWithAI = async (documentContent, documentName) => {
      const response = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${config.openAIApiKey}`
        },
        body: JSON.stringify({
          model: "gpt-4o-mini",
          messages: [
            {
              role: "system",
              content: "You are a helpful AI assistant that analyzes documents and provides key insights."
            },
            {
              role: "user",
              content: `Analyze the following document titled "${documentName}" and provide a summary, important insights, and any risks or action items.

${documentContent}`
            }
          ]
        })
      });
      if (!response.ok) {
        const data2 = await response.json().catch(() => ({}));
        throw new Error(data2.error?.message || "OpenAI request failed.");
      }
      const data = await response.json();
      return data.choices?.[0]?.message?.content || "No analysis returned.";
    };
    const runAnalysis = async () => {
      setLoading(true);
      setError("");
      setResults([]);
      try {
        const documents = await getDocumentsFromSharePoint();
        const processedDocuments = [];
        for (const doc of documents) {
          const content = await fetch(doc.downloadUrl).then((res) => res.text()).catch(() => "Document content could not be read.");
          const analysis = await analyzeDocumentWithAI(content, doc.name);
          processedDocuments.push({
            name: doc.name,
            analysis
          });
        }
        setResults(processedDocuments);
      } catch (err) {
        setError(err.message || "Something went wrong while analyzing documents.");
      } finally {
        setLoading(false);
      }
    };
    return /* @__PURE__ */ import_react.default.createElement("div", { style: { maxWidth: 900, margin: "40px auto", fontFamily: "Arial, sans-serif", padding: 24 } }, /* @__PURE__ */ import_react.default.createElement("h1", null, "SharePoint AI Agent"), /* @__PURE__ */ import_react.default.createElement("div", { style: { display: "grid", gap: 12, marginBottom: 20 } }, /* @__PURE__ */ import_react.default.createElement("label", null, "SharePoint site URL", /* @__PURE__ */ import_react.default.createElement(
      "input",
      {
        type: "text",
        name: "siteUrl",
        value: config.siteUrl,
        onChange: handleChange,
        style: inputStyle
      }
    )), /* @__PURE__ */ import_react.default.createElement("label", null, "Folder path", /* @__PURE__ */ import_react.default.createElement(
      "input",
      {
        type: "text",
        name: "folderPath",
        value: config.folderPath,
        onChange: handleChange,
        style: inputStyle
      }
    )), /* @__PURE__ */ import_react.default.createElement("label", null, "SharePoint username", /* @__PURE__ */ import_react.default.createElement(
      "input",
      {
        type: "text",
        name: "username",
        value: config.username,
        onChange: handleChange,
        style: inputStyle
      }
    )), /* @__PURE__ */ import_react.default.createElement("label", null, "SharePoint password", /* @__PURE__ */ import_react.default.createElement(
      "input",
      {
        type: "password",
        name: "password",
        value: config.password,
        onChange: handleChange,
        style: inputStyle
      }
    )), /* @__PURE__ */ import_react.default.createElement("label", null, "OpenAI API key", /* @__PURE__ */ import_react.default.createElement(
      "input",
      {
        type: "password",
        name: "openAIApiKey",
        value: config.openAIApiKey,
        onChange: handleChange,
        style: inputStyle
      }
    ))), /* @__PURE__ */ import_react.default.createElement(
      "button",
      {
        type: "button",
        onClick: runAnalysis,
        disabled: loading,
        style: {
          padding: "12px 20px",
          backgroundColor: "#2563eb",
          color: "#fff",
          border: "none",
          borderRadius: 8,
          cursor: loading ? "not-allowed" : "pointer"
        }
      },
      loading ? "Analyzing..." : "Analyze documents"
    ), error && /* @__PURE__ */ import_react.default.createElement("div", { style: { marginTop: 20, color: "crimson", background: "#fee2e2", padding: 12, borderRadius: 8 } }, error), /* @__PURE__ */ import_react.default.createElement("div", { style: { marginTop: 28 } }, results.length > 0 && /* @__PURE__ */ import_react.default.createElement(import_react.default.Fragment, null, /* @__PURE__ */ import_react.default.createElement("h2", null, "Analysis results"), results.map((item) => /* @__PURE__ */ import_react.default.createElement("div", { key: item.name, style: { marginBottom: 20, border: "1px solid #e5e7eb", borderRadius: 8, padding: 16 } }, /* @__PURE__ */ import_react.default.createElement("h3", null, item.name), /* @__PURE__ */ import_react.default.createElement("p", { style: { whiteSpace: "pre-wrap", margin: 0 } }, item.analysis))))));
  }
  var inputStyle = {
    width: "100%",
    padding: "10px 12px",
    marginTop: 6,
    border: "1px solid #d1d5db",
    borderRadius: 6,
    boxSizing: "border-box"
  };
})();
