import React, { useState } from 'react';

const initialConfig = {
  siteUrl: 'https://your-sharepoint-site-url/sites/your-site-name',
  folderPath: 'Shared Documents/Your-Folder-Name',
  openAIApiKey: '',
  username: '',
  password: '',
};

export default function App() {
  const [config, setConfig] = useState(initialConfig);
  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState([]);
  const [error, setError] = useState('');

  const handleChange = (event) => {
    const { name, value } = event.target;
    setConfig((prev) => ({ ...prev, [name]: value }));
  };

  const getDocumentsFromSharePoint = async () => {
    const response = await fetch('/api/sharepoint/documents', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        siteUrl: config.siteUrl,
        folderPath: config.folderPath,
        username: config.username,
        password: config.password,
      }),
    });

    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data.message || 'Unable to fetch documents from SharePoint.');
    }

    return response.json();
  };

  const analyzeDocumentWithAI = async (documentContent, documentName) => {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${config.openAIApiKey}`,
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [
          {
            role: 'system',
            content:
              'You are a helpful AI assistant that analyzes documents and provides key insights.',
          },
          {
            role: 'user',
            content: `Analyze the following document titled "${documentName}" and provide a summary, important insights, and any risks or action items.\n\n${documentContent}`,
          },
        ],
      }),
    });

    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data.error?.message || 'OpenAI request failed.');
    }

    const data = await response.json();
    return data.choices?.[0]?.message?.content || 'No analysis returned.';
  };

  const runAnalysis = async () => {
    setLoading(true);
    setError('');
    setResults([]);

    try {
      const documents = await getDocumentsFromSharePoint();
      const processedDocuments = [];

      for (const doc of documents) {
        const content = await fetch(doc.downloadUrl)
          .then((res) => res.text())
          .catch(() => 'Document content could not be read.');

        const analysis = await analyzeDocumentWithAI(content, doc.name);
        processedDocuments.push({
          name: doc.name,
          analysis,
        });
      }

      setResults(processedDocuments);
    } catch (err) {
      setError(err.message || 'Something went wrong while analyzing documents.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ maxWidth: 900, margin: '40px auto', fontFamily: 'Arial, sans-serif', padding: 24 }}>
      <h1>SharePoint AI Agent</h1>

      <div style={{ display: 'grid', gap: 12, marginBottom: 20 }}>
        <label>
          SharePoint site URL
          <input
            type="text"
            name="siteUrl"
            value={config.siteUrl}
            onChange={handleChange}
            style={inputStyle}
          />
        </label>

        <label>
          Folder path
          <input
            type="text"
            name="folderPath"
            value={config.folderPath}
            onChange={handleChange}
            style={inputStyle}
          />
        </label>

        <label>
          SharePoint username
          <input
            type="text"
            name="username"
            value={config.username}
            onChange={handleChange}
            style={inputStyle}
          />
        </label>

        <label>
          SharePoint password
          <input
            type="password"
            name="password"
            value={config.password}
            onChange={handleChange}
            style={inputStyle}
          />
        </label>

        <label>
          OpenAI API key
          <input
            type="password"
            name="openAIApiKey"
            value={config.openAIApiKey}
            onChange={handleChange}
            style={inputStyle}
          />
        </label>
      </div>

      <button
        type="button"
        onClick={runAnalysis}
        disabled={loading}
        style={{
          padding: '12px 20px',
          backgroundColor: '#2563eb',
          color: '#fff',
          border: 'none',
          borderRadius: 8,
          cursor: loading ? 'not-allowed' : 'pointer',
        }}
      >
        {loading ? 'Analyzing...' : 'Analyze documents'}
      </button>

      {error && (
        <div style={{ marginTop: 20, color: 'crimson', background: '#fee2e2', padding: 12, borderRadius: 8 }}>
          {error}
        </div>
      )}

      <div style={{ marginTop: 28 }}>
        {results.length > 0 && (
          <>
            <h2>Analysis results</h2>
            {results.map((item) => (
              <div key={item.name} style={{ marginBottom: 20, border: '1px solid #e5e7eb', borderRadius: 8, padding: 16 }}>
                <h3>{item.name}</h3>
                <p style={{ whiteSpace: 'pre-wrap', margin: 0 }}>{item.analysis}</p>
              </div>
            ))}
          </>
        )}
      </div>
    </div>
  );
}

const inputStyle = {
  width: '100%',
  padding: '10px 12px',
  marginTop: 6,
  border: '1px solid #d1d5db',
  borderRadius: 6,
  boxSizing: 'border-box',
};
