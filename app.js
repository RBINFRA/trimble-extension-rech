(function () {
  const PROPERTY_SET = "PSET - Attributs Mensura";
  const CRITERIA = [
    { id: "nom", label: "NOM" },
    { id: "zone", label: "ZONE" },
    { id: "source", label: "SOURCE" },
    { id: "entreprise", label: "ENTREPRISE D'EXÉCUTION" },
  ];

  const selectors = {
    nom: document.getElementById("nom"),
    zone: document.getElementById("zone"),
    source: document.getElementById("source"),
    entreprise: document.getElementById("entreprise"),
    dateMin: document.getElementById("dateMin"),
    dateMax: document.getElementById("dateMax"),
    searchBtn: document.getElementById("searchBtn"),
    resetBtn: document.getElementById("resetBtn"),
    status: document.getElementById("status"),
    error: document.getElementById("error"),
    resultCount: document.getElementById("resultCount"),
    summary: document.getElementById("summary"),
  };

  let API;

  function setStatus(message) {
    selectors.status.textContent = message || "";
  }

  function setError(message) {
    selectors.error.textContent = message || "";
  }

  function readCriteria() {
    const values = {};
    CRITERIA.forEach((c) => {
      const val = selectors[c.id].value.trim();
      if (val) values[c.label] = val;
    });
    const dateMin = selectors.dateMin.value;
    const dateMax = selectors.dateMax.value;
    return { values, dateMin: dateMin || undefined, dateMax: dateMax || undefined };
  }

  function buildPropertyFilters(values) {
    const props = {};
    Object.entries(values).forEach(([label, value]) => {
      props[`${PROPERTY_SET}.${label}`] = value;
    });
    return props;
  }

  function getPropertyValue(propertySets, targetName, propName) {
    if (!propertySets) return undefined;
    const pset = propertySets.find((p) => p.name === targetName);
    if (!pset || !pset.properties) return undefined;
    const prop = pset.properties.find((p) => p.name === propName);
    return prop ? prop.value : undefined;
  }

  function withinDateRange(value, start, end) {
    if (!value) return false;
    const parsed = new Date(value);
    if (isNaN(parsed.getTime())) return false;
    if (start && parsed < new Date(start)) return false;
    if (end && parsed > new Date(end)) return false;
    return true;
  }

  function updateSummary(matches, criteria) {
    selectors.summary.innerHTML = "";
    if (!matches.length) {
      selectors.resultCount.textContent = "Aucun élément trouvé.";
      return;
    }

    selectors.resultCount.textContent = `${matches.length} élément(s) trouvé(s).`;

    const addItem = (label, count) => {
      const li = document.createElement("li");
      li.innerHTML = `<span>${label}</span><span class="badge">${count}</span>`;
      selectors.summary.appendChild(li);
    };

    Object.entries(criteria.values).forEach(([label]) => {
      const count = matches.filter((m) => {
        const val = getPropertyValue(m.properties, PROPERTY_SET, label);
        if (val === undefined || val === null) return false;
        return String(val).toLowerCase().includes(criteria.values[label].toLowerCase());
      }).length;
      addItem(label, count);
    });

    if (criteria.dateMin || criteria.dateMax) {
      const count = matches.filter((m) => {
        const dateVal = getPropertyValue(m.properties, PROPERTY_SET, "DATE");
        return withinDateRange(dateVal, criteria.dateMin, criteria.dateMax);
      }).length;
      const label = criteria.dateMin && criteria.dateMax
        ? `DATE entre ${criteria.dateMin} et ${criteria.dateMax}`
        : criteria.dateMin
          ? `DATE après ${criteria.dateMin}`
          : `DATE avant ${criteria.dateMax}`;
      addItem(label, count);
    }
  }

  function buildSelectionPayload(matches) {
    const grouped = new Map();
    matches.forEach((m) => {
      const ids = grouped.get(m.modelId) || [];
      ids.push(m.id);
      grouped.set(m.modelId, ids);
    });

    const modelObjectIds = Array.from(grouped.entries()).map(([modelId, ids]) => ({
      modelId,
      objectRuntimeIds: ids,
    }));

    return { modelObjectIds };
  }

  async function highlightAndZoom(matches) {
    if (!API || !matches.length) return;
    const selector = buildSelectionPayload(matches);
    await API.viewer.setSelection(selector, "set");
    await API.viewer.setCamera(selector, { animationTime: 0.6 });
  }

  async function runSearch() {
    setError("");
    const criteria = readCriteria();
    if (!Object.keys(criteria.values).length && !criteria.dateMin && !criteria.dateMax) {
      setError("Veuillez saisir au moins un critère.");
      return;
    }

    setStatus("Connexion au viewer et récupération des objets...");

    try {
      const properties = buildPropertyFilters(criteria.values);
      const selector = Object.keys(properties).length
        ? { parameter: { properties } }
        : undefined;

      const objects = await API.viewer.getObjects(selector);
      const flattened = (objects || []).flatMap((m) =>
        (m.objects || []).map((o) => ({
          modelId: m.modelId,
          id: o.id,
          properties: o.properties,
        }))
      );

      const filtered = flattened.filter((obj) => {
        const dateValue = getPropertyValue(obj.properties, PROPERTY_SET, "DATE");
        const matchesDate = criteria.dateMin || criteria.dateMax
          ? withinDateRange(dateValue, criteria.dateMin, criteria.dateMax)
          : true;
        return matchesDate;
      });

      updateSummary(filtered, criteria);
      await highlightAndZoom(filtered);
      setStatus("Recherche terminée. Les éléments sont sélectionnés et zoomés.");
    } catch (err) {
      console.error(err);
      setError("Erreur lors de la recherche. Vérifiez la connexion au viewer.");
      setStatus("");
    }
  }

  function resetForm() {
    Object.values(selectors).forEach((el) => {
      if (el instanceof HTMLInputElement) el.value = "";
    });
    selectors.summary.innerHTML = "";
    selectors.resultCount.textContent = "Aucun résultat pour le moment.";
    setStatus("");
    setError("");
  }

  async function init() {
    setStatus("Connexion à Trimble Connect...");
    try {
      API = await TrimbleConnectWorkspace.connect(window.parent, (event, data) => {
        if (event === "extension.accessToken") {
          console.log("Token mis à jour", data);
        }
      }, 30000);
      setStatus("Connecté. Prêt pour la recherche.");
    } catch (err) {
      console.error(err);
      setError("Impossible de se connecter à Trimble Connect. Vérifiez l'extension.");
      setStatus("");
    }
  }

  selectors.searchBtn.addEventListener("click", runSearch);
  selectors.resetBtn.addEventListener("click", resetForm);

  window.addEventListener("DOMContentLoaded", init);
})();
