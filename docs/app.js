(function () {
  const $ = (id) => document.getElementById(id);
  const money = (n) => {
    if (!isFinite(n)) return "—";
    const abs = Math.abs(n);
    const sign = n < 0 ? "−" : "";
    if (abs >= 1e9) return sign + "$" + (abs / 1e9).toFixed(2) + "B";
    if (abs >= 1e6) return sign + "$" + (abs / 1e6).toFixed(2) + "M";
    if (abs >= 1e3) return sign + "$" + (abs / 1e3).toFixed(0) + "k";
    return sign + "$" + abs.toFixed(0);
  };

  function run() {
    const unique = Number($("unique").value) || 0;
    const nre = Number($("nre").value) || 0;
    const catalogItems = Number($("catalogItems").value) || 0;
    const unusedLb = Number($("unusedLb").value) || 0;
    const perLb = Number($("perLb").value) || 0;
    const moduleLb = Number($("moduleLb").value) || 0;
    const pLoss = Number($("pLoss").value) || 0;

    const designsAvoided = Math.max(0, unique - catalogItems);
    const nreSaved = designsAvoided * nre;
    const unusedSaved = unusedLb * perLb;
    const lossTraditional = (pLoss / 100) * moduleLb * perLb;
    const lossPacket = (pLoss / 100) * 1 * perLb;
    const lossSaved = Math.max(0, lossTraditional - lossPacket);
    const total = nreSaved + unusedSaved + lossSaved;

    $("nreSaved").textContent = money(nreSaved);
    $("unusedSaved").textContent = money(unusedSaved);
    $("lossSaved").textContent = money(lossSaved);
    $("totalSaved").textContent = money(total);
    $("designsAvoided").textContent = String(designsAvoided);
  }

  ["unique", "nre", "catalogItems", "unusedLb", "perLb", "moduleLb", "pLoss"].forEach((id) => {
    const el = $(id);
    if (el) el.addEventListener("input", run);
  });
  run();
})();
