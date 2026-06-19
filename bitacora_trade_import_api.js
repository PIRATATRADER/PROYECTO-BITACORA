(function () {
  const DEFAULT_API_URL = "https://us-central1-proyecto-bitacora-fead3.cloudfunctions.net/extractTradesFromImage";

  function getApiUrl() {
    return window.BITACORA_IMPORT_API_URL || DEFAULT_API_URL;
  }

  function toNumberOrNull(value) {
    if (value === null || value === undefined || value === "") return null;
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    const cleaned = String(value).replace(/[$,%Rr\s,]/g, "");
    if (!cleaned) return null;
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function normalizeTicker(value) {
    return String(value || "").trim().toUpperCase().replace(/\s+/g, "");
  }

  function normalizeDirection(value) {
    const raw = String(value || "").trim().toLowerCase();
    if (raw === "short" || raw === "sell short") return "Short";
    if (raw === "long" || raw === "margin" || raw === "buy") return "Long";
    return "Long";
  }

  function normalizeTrade(row) {
    const side = normalizeDirection(row.direccion || row.type || row.side);
    const realized = toNumberOrNull(row.realized ?? row.realized_pnl ?? row.pnl);
    const rValue = toNumberOrNull(row.realIR ?? row.r_multiple ?? row.rr);

    return {
      ticker: normalizeTicker(row.ticker || row.symbol),
      type: side === "Short" ? "Short" : "Margin",
      direccion: side,
      pnlShort: side === "Short" ? realized : null,
      rrShort: side === "Short" ? rValue : null,
      pnlLong: side === "Long" ? realized : null,
      rrLong: side === "Long" ? rValue : null,
      realized: realized,
      realIR: rValue,
      precioEntrada: row.precioEntrada ?? row.entry_price ?? null,
      setupEntrada: row.setupEntrada ?? row.setup ?? "",
      inicio: row.inicio ?? row.entry_time ?? "",
      closed: row.closed ?? row.exit_time ?? "",
      shares: row.shares ?? null,
      commissions: row.commissions ?? null,
      pais: row.pais ?? "",
      sector: row.sector ?? "",
      industry: row.industry ?? "",
      marketCap: row.marketCap ?? "",
      float: row.float ?? "",
      notas: row.notas ?? row.notes ?? "",
      confidence: row.confidence ?? null
    };
  }

  async function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result).split(",")[1]);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  async function extractFromBase64(payload) {
    const response = await fetch(getApiUrl(), {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data.error || `HTTP ${response.status}`);
    }

    const trades = Array.isArray(data.trades) ? data.trades.map(normalizeTrade).filter(row => row.ticker) : [];
    return {
      trades,
      provider: data.provider || "backend",
      model: data.model || "",
      warning: data.warning || "",
      confidence: data.confidence || "",
      rawCount: data.rawCount || trades.length
    };
  }

  async function extractFromFile(file, extra) {
    const imageBase64 = await fileToBase64(file);
    return extractFromBase64({
      imageBase64,
      imageMimeType: file.type || "image/png",
      imageName: file.name || "capture.png",
      tradeDate: extra && extra.tradeDate ? extra.tradeDate : "",
      source: extra && extra.source ? extra.source : "bitacora"
    });
  }

  window.BitacoraTradeImportApi = {
    extractFromBase64,
    extractFromFile,
    fileToBase64,
    normalizeTrade
  };
})();
