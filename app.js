(function () {
  const PROPERTY_SET = "PSET - Attributs Mensura";
  const CRITERIA = [
    { id: "element", label: "ELEMENTS" },
    { id: "localisation", label: "LOCALISATION" },
    { id: "nomProjet", label: "NOM PROJET" },
    { id: "source", label: "SOURCE" },
    { id: "entreprise", label: "ENTREPRISE D'EXÉCUTION" },
  ];
  const BATCH_SIZE = 200;

  const selectors = {
    element: document.getElementById("element"),
    localisation: document.getElementById("localisation"),
    nomProjet: document.getElementById("nomProjet"),
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
    progress: document.getElementById("loadingProgress"),
    progressBar: document.getElementById("loadingProgressBar"),
    progressText: document.getElementById("loadingProgressText"),
    loadingSpinner: document.getElementById("loadingSpinner"),
  };

  let API;
  let cachedObjects = [];
  let valueCatalog = {};
  let dataLoaded = false;
  let loadingPromise = null;
  const MAX_DATA_LOAD_ITERATIONS = 3;
  let progressPercent = 0;
  const collator = new Intl.Collator("fr", { sensitivity: "base" });
  const TYPE_PROP_STRAIGHT = "TYPE D'OBJET 3D";
  const TYPE_PROP_UNICODE_APOSTROPHE = "TYPE D\u2019OBJET 3D";
  const TYPE_PROPS = [TYPE_PROP_STRAIGHT, TYPE_PROP_UNICODE_APOSTROPHE];
  const ELEMENT_PROPS = ["ELEMENT", "ELEMENTS", "ÉLÉMENT", "ÉLÉMENTS"];
  const ELEMENT_FALLBACK = "Élément non spécifié";
  const SURFACE_PROP = "SURFACE";
  const LENGTH_PROP = "LONGUEUR";
  const TYPE_GROUPS = [
    { key: "SURFACIQUE", label: "TYPE D’OBJET 3D : SURFACIQUE", metric: { propNames: [SURFACE_PROP], label: "Surface totale", unit: "m²" } },
    { key: "LINÉAIRE", label: "TYPE D’OBJET 3D : LINÉAIRE", metric: { propNames: [LENGTH_PROP], label: "Longueur totale", unit: "ml" } },
    { key: "PONCTUEL", label: "TYPE D’OBJET 3D : PONCTUEL", metric: null },
  ];

  function normalizeModelState(state) {
    return typeof state === "string" ? state.toLowerCase() : undefined;
  }

  function setSelectVisualState(select) {
    if (!select) return;
    if (select.value) {
      select.classList.add("selected");
      select.classList.remove("placeholder");
    } else {
      select.classList.add("placeholder");
      select.classList.remove("selected");
    }
  }

  function bindSelectStateUpdates() {
    CRITERIA.forEach((c) => {
      const select = selectors[c.id];
      if (!select) return;
      select.addEventListener("change", () => setSelectVisualState(select));
      setSelectVisualState(select);
    });
  }

  function chunkArray(array, size) {
    const result = [];
    for (let i = 0; i < array.length; i += size) {
      result.push(array.slice(i, i + size));
    }
    return result;
  }

  function setStatus(message) {
    selectors.status.textContent = message || "";
  }

  function setError(message) {
    selectors.error.textContent = message || "";
  }

  function resetProgress() {
    if (!selectors.progress || !selectors.progressBar || !selectors.progressText) return;
    selectors.progress.classList.add("hidden");
    selectors.progressText.classList.add("hidden");
    selectors.progressBar.style.width = "0%";
    selectors.progressText.textContent = "";
    progressPercent = 0;
    if (selectors.loadingSpinner) {
      selectors.loadingSpinner.classList.add("hidden");
    }
  }

  function updateProgress(current, total) {
    if (!selectors.progress || !selectors.progressBar || !selectors.progressText) return;
    selectors.progress.classList.remove("hidden");
    selectors.progressText.classList.remove("hidden");
    if (selectors.loadingSpinner) {
      selectors.loadingSpinner.classList.remove("hidden");
    }
    const computed = total ? Math.min(100, Math.round((current / total) * 100)) : progressPercent;
    progressPercent = Math.max(progressPercent, computed);
    selectors.progressBar.style.width = `${progressPercent}%`;
    selectors.progressText.textContent = `Chargement des données... ${progressPercent}%`;
    return progressPercent;
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

  function getPropertyValue(propertySets, targetName, propName) {
    if (!propertySets) return undefined;
    const pset = propertySets.find((p) => equalsIgnoreCase(p.name, targetName));
    if (!pset || !pset.properties) return undefined;
    const prop = pset.properties.find((p) => equalsIgnoreCase(p.name, propName));
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

  function equalsIgnoreCase(a, b) {
    if (a === undefined || a === null || b === undefined || b === null) return false;
    return collator.compare(String(a).trim(), String(b).trim()) === 0;
  }

  function getPropertyValueWithFallback(propertySets, ...names) {
    for (const name of names) {
      const value = getPropertyValue(propertySets, PROPERTY_SET, name);
      if (value !== undefined && value !== null) return value;
    }
    return undefined;
  }

  function toNumericValue(value) {
    if (typeof value === "number") return value;
    if (value === undefined || value === null) return 0;
    const cleaned = String(value).replace(/\s/g, "").replace(",", ".");
    const parsed = parseFloat(cleaned);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 2 }).format(value);
  }

  function formatObjectCount(count) {
    return `${count} objet${count === 1 ? "" : "s"}`;
  }

  function aggregateByType(matches) {
    const map = new Map();
    let unknownCount = 0;
    matches.forEach((obj) => {
      const typeValue = getPropertyValueWithFallback(obj.properties, ...TYPE_PROPS);
      const group = TYPE_GROUPS.find((g) => equalsIgnoreCase(typeValue, g.key));
      if (!group) {
        unknownCount += 1;
        return;
      }
      const elementValue =
        getPropertyValueWithFallback(obj.properties, ...ELEMENT_PROPS) || ELEMENT_FALLBACK;
      const existingGroup = map.get(group.key) || { meta: group, items: new Map() };
      const existingItem = existingGroup.items.get(elementValue) || { element: elementValue, count: 0, total: 0 };
      existingItem.count += 1;
      if (group.metric) {
        const metricValue = getPropertyValueWithFallback(obj.properties, ...(group.metric.propNames || []));
        existingItem.total += toNumericValue(metricValue);
      }
      existingGroup.items.set(elementValue, existingItem);
      map.set(group.key, existingGroup);
    });

    const groups = TYPE_GROUPS.filter((g) => map.has(g.key)).map((g) => {
      const group = map.get(g.key);
      return { meta: group.meta, items: Array.from(group.items.values()) };
    });

    return { groups, unknownCount };
  }

  function buildValueCatalog(objects) {
    const catalog = {};
    CRITERIA.forEach((c) => (catalog[c.label] = new Set()));

    objects.forEach((obj) => {
      CRITERIA.forEach((c) => {
        const val = getPropertyValue(obj.properties, PROPERTY_SET, c.label);
        if (val === undefined || val === null) return;
        const normalized = String(val).trim(); // ignore empty / whitespace values
        if (normalized) catalog[c.label].add(normalized);
      });
    });

    return catalog;
  }

  function populateDropdowns(catalog) {
    CRITERIA.forEach((c) => {
      const select = selectors[c.id];
      if (!select) return;
      const currentValue = select.value;
      select.innerHTML = "";
      const empty = document.createElement("option");
      empty.value = "";
      empty.textContent = "-- Sélectionner --";
      select.appendChild(empty);

      const values = Array.from(catalog[c.label] || []).sort(collator.compare);
      values.forEach((val) => {
        const opt = document.createElement("option");
        opt.value = val;
        opt.textContent = val;
        select.appendChild(opt);
      });

      if (currentValue && values.includes(currentValue)) {
        select.value = currentValue;
      }
      setSelectVisualState(select);
    });
  }

  function flattenObjects(objects) {
    return (objects || []).flatMap((m) =>
      (m.objects || []).map((o) => ({
        modelId: m.modelId,
        id: o.id,
        properties: o.properties,
      }))
    );
  }

  function resetLoadedData(options = { keepForm: false }) {
    if (!options.keepForm) resetForm();
    cachedObjects = [];
    valueCatalog = {};
    dataLoaded = false;
    loadingPromise = null;
    resetProgress();
    populateDropdowns(valueCatalog);
  }

  async function fetchObjectsWithProperties(models, onProgress) {
    const result = [];
    const totalBatches = (models || []).reduce((sum, model) => {
      const ids = (model.objects || []).map((o) => o.id);
      return sum + Math.ceil(ids.length / BATCH_SIZE);
    }, 0);
    let processedBatches = 0;
    const report = () => {
      processedBatches += 1;
      if (typeof onProgress === "function") onProgress(processedBatches, totalBatches);
    };

    for (const model of models || []) {
      const ids = (model.objects || []).map((o) => o.id);
      if (!ids.length) {
        result.push({ modelId: model.modelId, objects: [] });
        continue;
      }

      const batches = chunkArray(ids, BATCH_SIZE);
      const objects = [];
      for (const batch of batches) {
        const props = await API.viewer.getObjectProperties(model.modelId, batch);
        objects.push(...props);
        report();
      }

      result.push({ modelId: model.modelId, objects });
    }
    if (!totalBatches && typeof onProgress === "function") {
      onProgress(0, 0);
    }
    return result;
  }

  async function ensureDataLoaded() {
    let retryCount = 0;
    while (retryCount < MAX_DATA_LOAD_ITERATIONS) {
      if (dataLoaded) return;
      if (loadingPromise) {
        await loadingPromise;
        if (dataLoaded) return;
      } else {
        setStatus("Récupération des données disponibles...");
        updateProgress(0, 0);
        const inFlight = (async () => {
          const models = await API.viewer.getObjects();
          const objectsWithProps = await fetchObjectsWithProperties(models, (current, total) => {
            const percent = updateProgress(current, total);
            const status = total
              ? `Chargement des données... ${percent}%`
              : "Chargement des données...";
            setStatus(status);
          });
          cachedObjects = flattenObjects(objectsWithProps);
          valueCatalog = buildValueCatalog(cachedObjects);
          populateDropdowns(valueCatalog);
          dataLoaded = true;
          setStatus("Données chargées. Prêt pour la recherche.");
          resetProgress();
        })();
        loadingPromise = inFlight;

        try {
          await inFlight;
        } catch (err) {
          resetLoadedData({ keepForm: true });
          console.error(err);
          setError("Impossible de récupérer les propriétés des objets. Vérifiez le chargement du modèle.");
          setStatus("");
        } finally {
          if (loadingPromise === inFlight) {
            loadingPromise = null;
          }
        }
      }

      if (dataLoaded) return;
      retryCount += 1;
      // If we reach here, data is still absent (e.g., a model unload/reset event cleared caches during load); retry while iterations remain.
    }
  }

  function matchesAllCriteria(obj, criteria) {
    const matchesProperties = Object.entries(criteria.values).every(([label, value]) => {
      const propVal = getPropertyValue(obj.properties, PROPERTY_SET, label);
      if (propVal === undefined || propVal === null) return false;
      return equalsIgnoreCase(propVal, value);
    });

    if (!matchesProperties) return false;

    if (criteria.dateMin || criteria.dateMax) {
      const dateValue = getPropertyValue(obj.properties, PROPERTY_SET, "DATE");
      return withinDateRange(dateValue, criteria.dateMin, criteria.dateMax);
    }

    return true;
  }

  function updateSummary(matches) {
    selectors.summary.innerHTML = "";
    if (!matches.length) {
      selectors.resultCount.textContent = "Aucun élément trouvé.";
      return;
    }

    selectors.resultCount.textContent = `${formatObjectCount(matches.length)} sélectionné(s).`;

    const { groups, unknownCount } = aggregateByType(matches);
    if (!groups.length && !unknownCount) {
      const li = document.createElement("li");
      li.textContent = "Aucun regroupement disponible (TYPE D'OBJET 3D manquant).";
      selectors.summary.appendChild(li);
      return;
    }

    groups.forEach((group) => {
      group.items.forEach((item) => {
        const li = document.createElement("li");
        const countLabel = formatObjectCount(item.count);
        const metricText = group.meta.metric
          ? ` – ${group.meta.metric.label} : ${formatNumber(item.total)}${group.meta.metric.unit ? ` ${group.meta.metric.unit}` : ""}`
          : "";
        li.innerHTML = `<span class="summary-item-text">${countLabel} - ${item.element}${metricText}</span>`;
        selectors.summary.appendChild(li);
      });
    });

    if (unknownCount > 0) {
      const li = document.createElement("li");
      li.textContent = `${formatObjectCount(unknownCount)} - INCONNUS`;
      selectors.summary.appendChild(li);
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
    await ensureDataLoaded();
    const criteria = readCriteria();
    if (!Object.keys(criteria.values).length && !criteria.dateMin && !criteria.dateMax) {
      setError("Veuillez saisir au moins un critère.");
      return;
    }

    setStatus("Application du filtrage en cours...");

    try {
      const filtered = cachedObjects.filter((obj) => matchesAllCriteria(obj, criteria));

      updateSummary(filtered);
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
      if (el instanceof HTMLInputElement || el instanceof HTMLSelectElement) el.value = "";
    });
    selectors.summary.innerHTML = "";
    selectors.resultCount.textContent = "Aucun résultat pour le moment.";
    setStatus("");
    setError("");
    CRITERIA.forEach((c) => setSelectVisualState(selectors[c.id]));
  }

  async function init() {
    setStatus("Connexion à Trimble Connect...");
    try {
      API = await TrimbleConnectWorkspace.connect(window.parent, async (event, data) => {
        if (event === "extension.accessToken") {
          console.log("Token mis à jour", data);
          return;
        }
        if (event === "embed.pageLoaded") {
          resetLoadedData();
          setStatus("Maquette rechargée. Mise à jour des filtres...");
          await ensureDataLoaded();
          return;
        }
        if (event === "viewer.onModelStateChanged") {
          const state = data?.data?.state;
          const normalizedState = normalizeModelState(state);
          if (normalizedState === "unloaded") {
            resetLoadedData();
            setStatus("Modèle déchargé. En attente d'un nouveau chargement...");
            return;
          }
          if (normalizedState === "loaded") {
            resetLoadedData();
            setStatus("Nouveau modèle chargé. Mise à jour des filtres...");
            await ensureDataLoaded();
            return;
          }
        }
        if (event === "viewer.onModelReset") {
          resetLoadedData();
          setStatus("Modèle réinitialisé. Mise à jour des filtres...");
          await ensureDataLoaded();
        }
      }, 30000);
      await ensureDataLoaded();
    } catch (err) {
      console.error(err);
      setError("Impossible de se connecter à Trimble Connect. Vérifiez l'extension.");
      setStatus("");
    }
  }

  selectors.searchBtn.addEventListener("click", runSearch);
  selectors.resetBtn.addEventListener("click", resetForm);
  bindSelectStateUpdates();

  window.addEventListener("DOMContentLoaded", init);
})();
