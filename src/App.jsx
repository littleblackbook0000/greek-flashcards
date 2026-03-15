import React, { useEffect, useMemo, useState } from "react";

const STORAGE_KEY = "greek-flashcards-simple-v3";

const starterDeck = [
  {
    id: crypto.randomUUID(),
    greek: "ἀγάπη",
    fullWords: "ἀγάπη, -ης, ἡ",
    meaning: "love",
    type: "Noun",
    frequency: 116,
    difficulty: 5,
    known: 0,
    missed: 0,
  },
  {
    id: crypto.randomUUID(),
    greek: "λόγος",
    fullWords: "λόγος, -ου, ὁ",
    meaning: "word, message",
    type: "Noun",
    frequency: 330,
    difficulty: 1,
    known: 0,
    missed: 0,
  },
  {
    id: crypto.randomUUID(),
    greek: "γινώσκω",
    fullWords: "γινώσκω",
    meaning: "to know",
    type: "Verb",
    frequency: 222,
    difficulty: 3,
    known: 0,
    missed: 0,
  },
];

function normalizeRow(row) {
  const greek = row.Word ?? row.word ?? row.Greek ?? row.greek ?? "";
  const fullWords = row["Full words"] ?? row.fullWords ?? row["Full Words"] ?? "";
  const meaning = row["Meaning (English Gloss)"] ?? row.meaning ?? row.Meaning ?? "";
  const type = row.Type ?? row.type ?? "";
  const frequency = Number(row.Frequency ?? row.frequency ?? 0) || 0;
  const difficulty = Number(row.Difficulty ?? row.difficulty ?? 0) || 0;

  if (!String(greek).trim() || !String(meaning).trim()) return null;

  return {
    id: crypto.randomUUID(),
    greek: String(greek).trim(),
    fullWords: String(fullWords || "").trim(),
    meaning: String(meaning).trim(),
    type: String(type || "").trim(),
    frequency,
    difficulty,
    known: 0,
    missed: 0,
  };
}

async function readExcelFile(file) {
  const XLSX = await import("xlsx");
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheet];
  return XLSX.utils.sheet_to_json(worksheet, { defval: "" });
}

export default function GreekVocabFlashcardsApp() {
  const [cards, setCards] = useState(starterDeck);
  const [selectedTypes, setSelectedTypes] = useState([]);
  const [freqMin, setFreqMin] = useState("");
  const [freqMax, setFreqMax] = useState("");
  const [diffMin, setDiffMin] = useState("");
  const [diffMax, setDiffMax] = useState("");
  const [index, setIndex] = useState(0);
  const [flipped, setFlipped] = useState(false);
  const [importMessage, setImportMessage] = useState("");

  useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return;
    try {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length) setCards(parsed);
    } catch {
      // ignore malformed local data
    }
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cards));
  }, [cards]);

  const availableTypes = useMemo(() => {
    return [...new Set(cards.map((c) => c.type).filter(Boolean))].sort();
  }, [cards]);

  const filteredCards = useMemo(() => {
    return cards.filter((card) => {
      const typeOk = selectedTypes.length === 0 || selectedTypes.includes(card.type);
      const freqOk =
        (freqMin === "" || Number(card.frequency) >= Number(freqMin)) &&
        (freqMax === "" || Number(card.frequency) <= Number(freqMax));
      const diffOk =
        (diffMin === "" || Number(card.difficulty) >= Number(diffMin)) &&
        (diffMax === "" || Number(card.difficulty) <= Number(diffMax));
      return typeOk && freqOk && diffOk;
    });
  }, [cards, selectedTypes, freqMin, freqMax, diffMin, diffMax]);

  const currentCard = useMemo(() => filteredCards[index] || null, [filteredCards, index]);
  const progress = filteredCards.length ? Math.round(((index + 1) / filteredCards.length) * 100) : 0;
  const score = useMemo(() => {
    const known = cards.reduce((sum, c) => sum + c.known, 0);
    const wrong = cards.reduce((sum, c) => sum + c.missed, 0);
    return { known, wrong };
  }, [cards]);

  useEffect(() => {
    setIndex(0);
    setFlipped(false);
  }, [selectedTypes, freqMin, freqMax, diffMin, diffMax]);

  useEffect(() => {
    if (index > filteredCards.length - 1) {
      setIndex(0);
      setFlipped(false);
    }
  }, [filteredCards.length, index]);

  function nextCard() {
    if (!filteredCards.length) return;
    setIndex((prev) => (prev + 1) % filteredCards.length);
    setFlipped(false);
  }

  function markCard(result) {
    if (!currentCard) return;
    setCards((prev) =>
      prev.map((c) =>
        c.id === currentCard.id
          ? {
              ...c,
              known: result === "correct" ? c.known + 1 : c.known,
              missed: result === "wrong" ? c.missed + 1 : c.missed,
            }
          : c
      )
    );
    nextCard();
  }

  function toggleType(type) {
    setSelectedTypes((prev) =>
      prev.includes(type) ? prev.filter((t) => t !== type) : [...prev, type]
    );
  }

  function clearFilters() {
    setSelectedTypes([]);
    setFreqMin("");
    setFreqMax("");
    setDiffMin("");
    setDiffMax("");
  }

  async function handleImport(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const rows = await readExcelFile(file);
      const imported = rows.map(normalizeRow).filter(Boolean);
      if (!imported.length) {
        setImportMessage("No valid vocabulary rows were found.");
        return;
      }
      setCards(imported);
      setSelectedTypes([]);
      setFreqMin("");
      setFreqMax("");
      setDiffMin("");
      setDiffMax("");
      setIndex(0);
      setFlipped(false);
      setImportMessage(`Imported ${imported.length} words from ${file.name}`);
    } catch (error) {
      setImportMessage("Import failed. Please check the Excel format.");
      console.error(error);
    } finally {
      event.target.value = "";
    }
  }

  return (
    <div style={styles.page}>
      <div style={styles.layout}>
        <div style={styles.shell}>
          <div style={styles.headerRow}>
            <div>
              <div style={styles.title}>Greek Flash Cards</div>
              <div style={styles.subtitle}>Simple study mode</div>
            </div>
            <div style={styles.counter}>{filteredCards.length ? index + 1 : 0} / {filteredCards.length}</div>
          </div>

          <div style={styles.progressTrack}>
            <div style={{ ...styles.progressFill, width: `${progress}%` }} />
          </div>

          {currentCard ? (
            <div
              style={{
                ...styles.card,
                background: flipped
                  ? "linear-gradient(135deg, #7c3aed 0%, #ec4899 100%)"
                  : "linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%)",
              }}
            >
              {!flipped ? (
                <>
                  <div style={styles.cardLabel}>Greek</div>
                  <div style={styles.greekText}>{currentCard.greek}</div>
                  <div style={styles.cardHint}>Tap flip to reveal the answer</div>
                </>
              ) : (
                <>
                  <div style={styles.cardLabel}>Answer</div>
                  <div style={styles.meaningText}>{currentCard.meaning}</div>
                  <div style={styles.answerMeta}>{currentCard.fullWords || currentCard.type || "Vocabulary"}</div>
                  <div style={styles.answerMetaSmall}>
                    {currentCard.type ? `Type ${currentCard.type}` : ""}
                    {currentCard.frequency ? `  •  Frequency ${currentCard.frequency}` : ""}
                    {currentCard.difficulty ? `  •  Difficulty ${currentCard.difficulty}` : ""}
                  </div>
                </>
              )}
            </div>
          ) : (
            <div style={styles.emptyCard}>No cards match the current filter.</div>
          )}

          <div style={styles.centerRow}>
            <button style={styles.flipButton} onClick={() => setFlipped((v) => !v)} disabled={!currentCard}>
              Flip
            </button>
          </div>

          {flipped && currentCard ? (
            <div style={styles.answerButtons}>
              <button style={styles.correctButton} onClick={() => markCard("correct")}>Correct</button>
              <button style={styles.wrongButton} onClick={() => markCard("wrong")}>Wrong</button>
            </div>
          ) : null}

          <div style={styles.footerRow}>
            <div style={styles.scoreBox}>Correct {score.known}</div>
            <div style={styles.scoreBox}>Wrong {score.wrong}</div>
          </div>

          <div style={styles.bottomBar}>
            <label style={styles.importButton}>
              Import
              <input type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleImport} />
            </label>
          </div>

          {importMessage ? <div style={styles.importMessage}>{importMessage}</div> : null}

          <div style={styles.helpBox}>
            Excel columns supported: Type, Word, Full words, Meaning (English Gloss), Frequency, Difficulty
          </div>
        </div>

        <div style={styles.sidebar}>
          <div style={styles.sidebarTitle}>Filters</div>
          <button style={styles.clearButton} onClick={clearFilters}>Clear all</button>

          <div style={styles.filterBlock}>
            <div style={styles.filterLabel}>Type</div>
            <div style={styles.typeList}>
              {availableTypes.map((type) => (
                <button
                  key={type}
                  style={{ ...styles.typeChip, ...(selectedTypes.includes(type) ? styles.typeChipActive : {}) }}
                  onClick={() => toggleType(type)}
                >
                  {type}
                </button>
              ))}
            </div>
          </div>

          <div style={styles.filterBlock}>
            <div style={styles.filterLabel}>Frequency</div>
            <div style={styles.rangeRow}>
              <input style={styles.rangeInput} value={freqMin} onChange={(e) => setFreqMin(e.target.value)} placeholder="Min" />
              <input style={styles.rangeInput} value={freqMax} onChange={(e) => setFreqMax(e.target.value)} placeholder="Max" />
            </div>
          </div>

          <div style={styles.filterBlock}>
            <div style={styles.filterLabel}>Difficulty</div>
            <div style={styles.rangeRow}>
              <input style={styles.rangeInput} value={diffMin} onChange={(e) => setDiffMin(e.target.value)} placeholder="Min" />
              <input style={styles.rangeInput} value={diffMax} onChange={(e) => setDiffMax(e.target.value)} placeholder="Max" />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

const styles = {
  page: {
    minHeight: "100vh",
    display: "flex",
    justifyContent: "center",
    alignItems: "flex-start",
    padding: 24,
    background: "linear-gradient(180deg, #fdf2f8 0%, #eef2ff 50%, #ecfeff 100%)",
    fontFamily: "Arial, Helvetica, sans-serif",
  },
  layout: {
    width: "100%",
    maxWidth: 1180,
    display: "grid",
    gridTemplateColumns: "minmax(0, 1fr) 300px",
    gap: 24,
    alignItems: "start",
  },
  shell: {
    width: "100%",
    background: "rgba(255,255,255,0.88)",
    backdropFilter: "blur(10px)",
    borderRadius: 28,
    padding: 24,
    boxShadow: "0 20px 60px rgba(79, 70, 229, 0.18)",
  },
  sidebar: {
    background: "rgba(255,255,255,0.92)",
    backdropFilter: "blur(10px)",
    borderRadius: 28,
    padding: 22,
    boxShadow: "0 18px 50px rgba(124, 58, 237, 0.14)",
    position: "sticky",
    top: 24,
  },
  sidebarTitle: {
    fontSize: 26,
    fontWeight: 800,
    color: "#312e81",
    marginBottom: 14,
  },
  clearButton: {
    width: "100%",
    border: "none",
    borderRadius: 16,
    padding: "12px 16px",
    fontSize: 16,
    fontWeight: 800,
    color: "#5b21b6",
    cursor: "pointer",
    background: "#ede9fe",
    marginBottom: 18,
  },
  filterBlock: {
    marginBottom: 22,
  },
  filterLabel: {
    fontSize: 16,
    fontWeight: 800,
    color: "#4338ca",
    marginBottom: 10,
  },
  typeList: {
    display: "flex",
    flexWrap: "wrap",
    gap: 8,
  },
  typeChip: {
    border: "none",
    borderRadius: 999,
    padding: "10px 14px",
    fontSize: 14,
    fontWeight: 700,
    color: "#334155",
    cursor: "pointer",
    background: "#f1f5f9",
  },
  typeChipActive: {
    color: "white",
    background: "linear-gradient(90deg, #2563eb 0%, #7c3aed 100%)",
  },
  rangeRow: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 10,
  },
  rangeInput: {
    width: "100%",
    border: "1px solid #d8b4fe",
    outline: "none",
    borderRadius: 14,
    padding: "12px 12px",
    fontSize: 15,
    background: "white",
  },
  headerRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 12,
    marginBottom: 18,
  },
  title: {
    fontSize: 32,
    fontWeight: 800,
    color: "#312e81",
  },
  subtitle: {
    fontSize: 15,
    color: "#7c3aed",
    marginTop: 4,
  },
  counter: {
    background: "#ede9fe",
    color: "#5b21b6",
    padding: "10px 14px",
    borderRadius: 999,
    fontWeight: 700,
  },
  progressTrack: {
    width: "100%",
    height: 12,
    borderRadius: 999,
    background: "#e5e7eb",
    overflow: "hidden",
    marginBottom: 22,
  },
  progressFill: {
    height: "100%",
    borderRadius: 999,
    background: "linear-gradient(90deg, #06b6d4 0%, #8b5cf6 50%, #ec4899 100%)",
  },
  card: {
    minHeight: 330,
    borderRadius: 28,
    color: "white",
    display: "flex",
    flexDirection: "column",
    justifyContent: "center",
    alignItems: "center",
    textAlign: "center",
    padding: 28,
    boxShadow: "0 20px 40px rgba(59, 130, 246, 0.18)",
  },
  cardLabel: {
    fontSize: 16,
    fontWeight: 700,
    opacity: 0.95,
    marginBottom: 16,
    letterSpacing: 0.5,
  },
  greekText: {
    fontSize: 54,
    fontWeight: 800,
    lineHeight: 1.2,
    fontFamily: "Noto Serif, Times New Roman, serif",
  },
  meaningText: {
    fontSize: 38,
    fontWeight: 800,
    lineHeight: 1.25,
  },
  answerMeta: {
    marginTop: 16,
    fontSize: 22,
    opacity: 0.95,
  },
  answerMetaSmall: {
    marginTop: 12,
    fontSize: 15,
    opacity: 0.9,
    lineHeight: 1.6,
  },
  cardHint: {
    marginTop: 18,
    fontSize: 16,
    opacity: 0.92,
  },
  centerRow: {
    display: "flex",
    justifyContent: "center",
    marginTop: 22,
  },
  flipButton: {
    border: "none",
    borderRadius: 999,
    padding: "16px 34px",
    fontSize: 20,
    fontWeight: 800,
    color: "white",
    cursor: "pointer",
    background: "linear-gradient(90deg, #f59e0b 0%, #ef4444 100%)",
    boxShadow: "0 12px 24px rgba(239, 68, 68, 0.2)",
  },
  answerButtons: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 14,
    marginTop: 18,
  },
  correctButton: {
    border: "none",
    borderRadius: 20,
    padding: "16px 18px",
    fontSize: 20,
    fontWeight: 800,
    color: "white",
    cursor: "pointer",
    background: "linear-gradient(90deg, #10b981 0%, #22c55e 100%)",
  },
  wrongButton: {
    border: "none",
    borderRadius: 20,
    padding: "16px 18px",
    fontSize: 20,
    fontWeight: 800,
    color: "white",
    cursor: "pointer",
    background: "linear-gradient(90deg, #f43f5e 0%, #ec4899 100%)",
  },
  footerRow: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 12,
    marginTop: 18,
  },
  scoreBox: {
    background: "#faf5ff",
    color: "#6d28d9",
    padding: 14,
    borderRadius: 18,
    textAlign: "center",
    fontWeight: 800,
    fontSize: 18,
  },
  bottomBar: {
    display: "flex",
    justifyContent: "center",
    marginTop: 24,
  },
  importButton: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: 999,
    padding: "14px 28px",
    fontSize: 18,
    fontWeight: 800,
    color: "white",
    cursor: "pointer",
    background: "linear-gradient(90deg, #2563eb 0%, #7c3aed 100%)",
    boxShadow: "0 12px 24px rgba(124, 58, 237, 0.2)",
  },
  importMessage: {
    marginTop: 14,
    textAlign: "center",
    color: "#334155",
    fontWeight: 700,
  },
  helpBox: {
    marginTop: 18,
    background: "#eff6ff",
    color: "#1d4ed8",
    borderRadius: 18,
    padding: 14,
    textAlign: "center",
    lineHeight: 1.6,
    fontSize: 14,
  },
  emptyCard: {
    minHeight: 220,
    background: "white",
    borderRadius: 24,
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    textAlign: "center",
    padding: 24,
    fontSize: 24,
    fontWeight: 700,
    color: "#334155",
    boxShadow: "0 12px 30px rgba(0,0,0,0.08)",
  },
};
