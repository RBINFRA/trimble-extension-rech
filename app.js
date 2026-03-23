(function () {
  const PROPERTY_SET = "PSET - Attributs Mensura";
  const CRITERIA = [
    { id: "element", label: "ELEMENTS" },
    { id: "localisation", label: "LOCALISATION" },
    { id: "nomProjet", label: "NOM PROJET" },
    { id: "source", label: "SOURCE" },
    { id: "entreprise", label: "ENTREPRISE D'EXÉCUTION" },
  ];

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
  };

  let API;
  let cachedObjects = [];
  let valueCatalog = {};
  let dataLoaded = false;
  let loadingPromise = null;
  const MAX_DATA_LOAD_ATTEMPTS = 3;
  const collator = new Intl.Collator("fr", { sensitivity: "base" });

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
    populateDropdowns(valueCatalog);
  }

  async function fetchObjectsWithProperties(models) {
    const result = [];
    for (const model of models || []) {
      const ids = (model.objects || []).map((o) => o.id);
      if (!ids.length) {
        result.push({ modelId: model.modelId, objects: [] });
        continue;
      }

      const batches = chunkArray(ids, 200);
      const objects = [];
      for (const batch of batches) {
        const props = await API.viewer.getObjectProperties(model.modelId, batch);
        objects.push(...props);
      }

      result.push({ modelId: model.modelId, objects });
    }
    return result;
  }

  async function ensureDataLoaded() {
    let loadAttempt = 0;
    while (loadAttempt < MAX_DATA_LOAD_ATTEMPTS) {
      if (dataLoaded) return;
      if (loadingPromise) {
        await loadingPromise;
        if (dataLoaded) return;
        continue;
      }

      setStatus("Récupération des données disponibles...");
      const inFlight = (async () => {
        const models = await API.viewer.getObjects();
        const objectsWithProps = await fetchObjectsWithProperties(models);
        cachedObjects = flattenObjects(objectsWithProps);
        valueCatalog = buildValueCatalog(cachedObjects);
        populateDropdowns(valueCatalog);
        dataLoaded = true;
        setStatus("Données chargées. Prêt pour la recherche.");
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

      if (dataLoaded) return;
      loadAttempt += 1;
      // If we reach here, data is still absent (possible reset during loading); retry while attempts remain.
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
        return equalsIgnoreCase(val, criteria.values[label]);
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
    await ensureDataLoaded();
    const criteria = readCriteria();
    if (!Object.keys(criteria.values).length && !criteria.dateMin && !criteria.dateMax) {
      setError("Veuillez saisir au moins un critère.");
      return;
    }

    setStatus("Application du filtrage en cours...");

    try {
      const filtered = cachedObjects.filter((obj) => matchesAllCriteria(obj, criteria));

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
          const normalizedState = typeof state === "string" ? state.toLowerCase() : undefined;
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
