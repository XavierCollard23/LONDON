const snowContainer = document.querySelector(".snow");
const lanternContainer = document.querySelector(".lanterns");
const navToggle = document.querySelector(".nav-toggle");
const navLinks = document.querySelector(".nav-links");
const programmeContainer = document.getElementById("programme-content");
const mapGallery = document.getElementById("map-gallery");

const escapeHtml = (unsafe) =>
  unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

const createSnowflakes = () => {
  if (!snowContainer) return;
  const flakes = window.innerWidth < 600 ? 40 : 70;
  const fragment = document.createDocumentFragment();
  for (let i = 0; i < flakes; i += 1) {
    const flake = document.createElement("span");
    const size = Math.random() * 4 + 2;
    const duration = Math.random() * 15 + 12;
    const delay = Math.random() * -20;
    flake.style.setProperty("--x-start", `${Math.random() * 100}vw`);
    flake.style.setProperty("--x-end", `${Math.random() * 100 - 50}vw`);
    flake.style.animation = `snowfall ${duration}s linear ${delay}s infinite`;
    flake.style.position = "absolute";
    flake.style.left = `${Math.random() * 100}vw`;
    flake.style.top = `${Math.random() * 100 - 10}vh`;
    flake.style.width = `${size}px`;
    flake.style.height = `${size}px`;
    flake.style.borderRadius = "50%";
    flake.style.background = "rgba(255, 255, 255, 0.75)";
    flake.style.boxShadow = "0 0 8px rgba(255,255,255,0.3)";
    fragment.appendChild(flake);
  }
  snowContainer.appendChild(fragment);
};

const createLanterns = () => {
  if (!lanternContainer) return;
  const fragment = document.createDocumentFragment();
  const count = window.innerWidth < 600 ? 12 : 22;
  for (let i = 0; i < count; i += 1) {
    const lantern = document.createElement("span");
    lantern.className = "lantern";
    const duration = (Math.random() * 14 + 18).toFixed(2);
    const delay = (Math.random() * -24).toFixed(2);
    const scale = (Math.random() * 0.6 + 0.7).toFixed(2);
    const xStart = (Math.random() * 60 - 30).toFixed(2);
    const xEnd = (Math.random() * 60 - 30).toFixed(2);
    lantern.style.left = `${Math.random() * 100}vw`;
    lantern.style.setProperty("--duration", `${duration}s`);
    lantern.style.setProperty("--delay", `${delay}s`);
    lantern.style.setProperty("--scale", scale);
    lantern.style.setProperty("--x-start", `${xStart}vw`);
    lantern.style.setProperty("--x-end", `${xEnd}vw`);
    fragment.appendChild(lantern);
  }
  lanternContainer.appendChild(fragment);
};

const buildTimeline = (schedule) => {
  const fragment = document.createDocumentFragment();
  schedule.forEach((slot) => {
    const wrapper = document.createElement("div");
    wrapper.className = "slot";
    wrapper.innerHTML = `
      <div class="time">${escapeHtml(slot.time)}</div>
      <div class="slot-details">
        <p class="moment">${escapeHtml(slot.moment)}</p>
        <p class="ambiance">${escapeHtml(slot.ambiance)}</p>
      </div>
    `;
    fragment.appendChild(wrapper);
  });
  return fragment;
};

const buildMapTrigger = (index) => {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "map-trigger";
  button.setAttribute("aria-expanded", "false");
  button.setAttribute("aria-controls", `map-${index}`);
  button.innerHTML = `
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M4 6.5 9 4l6 2.5 5-2.5v13l-5 2.5-6-2.5-5 2.5zM9 6.5v11l6 2.5v-11z" />
    </svg>
    Voir la carte
  `;
  return button;
};

const renderProgramme = (days) => {
  if (!programmeContainer) return;
  programmeContainer.innerHTML = "";
  const fragment = document.createDocumentFragment();
  days.forEach((day, idx) => {
    const article = document.createElement("article");
    article.className = "day-card";
    article.id = `jour-${idx + 1}`;

    const header = document.createElement("div");
    header.className = "day-header";
    const number = document.createElement("p");
    number.className = "day-number";
    number.textContent = `Jour ${idx + 1}`;
    const title = document.createElement("h3");
    title.textContent = day.title;
    header.appendChild(number);
    header.appendChild(title);
    if (day.note) {
      const note = document.createElement("p");
      note.className = "day-note";
      note.textContent = day.note;
      header.appendChild(note);
    }
    article.appendChild(header);

    const timeline = document.createElement("div");
    timeline.className = "timeline";
    timeline.appendChild(buildTimeline(day.schedule));
    article.appendChild(timeline);

    const trigger = buildMapTrigger(idx + 1);
    article.appendChild(trigger);

    const mapWrapper = document.createElement("div");
    mapWrapper.className = "map-wrapper";
    mapWrapper.id = `map-${idx + 1}`;
    const iframe = document.createElement("iframe");
    iframe.title = `Carte du ${day.title}`;
    iframe.src = `maps/${day.map}`;
    mapWrapper.appendChild(iframe);
    article.appendChild(mapWrapper);

    trigger.addEventListener("click", () => {
      const isOpen = trigger.getAttribute("aria-expanded") === "true";
      trigger.setAttribute("aria-expanded", String(!isOpen));
      if (isOpen) {
        mapWrapper.classList.remove("expanded");
      } else {
        mapWrapper.classList.add("expanded");
      }
    });

    fragment.appendChild(article);
  });
  programmeContainer.appendChild(fragment);
};

const renderGallery = (days) => {
  if (!mapGallery) return;
  mapGallery.innerHTML = "";
  const fragment = document.createDocumentFragment();
  days.forEach((day, idx) => {
    const card = document.createElement("article");
    card.className = "map-card";
    const title = document.createElement("h3");
    title.textContent = `Jour ${idx + 1} — ${day.title}`;
    const description = document.createElement("p");
    description.textContent = day.note
      ? day.note
      : "Zoomer ou cliquer sur les marqueurs pour visualiser chaque étape de la journée.";
    const iframe = document.createElement("iframe");
    iframe.src = `maps/${day.map}`;
    iframe.title = `Itinéraire interactif du ${day.title}`;
    card.appendChild(title);
    card.appendChild(iframe);
    card.appendChild(description);
    fragment.appendChild(card);
  });
  mapGallery.appendChild(fragment);
};

const getInlineItinerary = () => {
  const node = document.getElementById("itinerary-data");
  if (!node) return null;
  try {
    return JSON.parse(node.textContent);
  } catch (error) {
    console.error("Impossible d'interpréter le programme embarqué.", error);
    return null;
  }
};

const loadItinerary = async () => {
  const inlineData = getInlineItinerary();
  if (inlineData && Array.isArray(inlineData)) {
    renderProgramme(inlineData);
    renderGallery(inlineData);
    return;
  }

  try {
    const response = await fetch("itinerary.json");
    if (!response.ok) {
      throw new Error(`Impossible de charger le programme (${response.status})`);
    }
    const days = await response.json();
    renderProgramme(days);
    renderGallery(days);
  } catch (error) {
    if (programmeContainer) {
      programmeContainer.innerHTML = `<p role="alert">Le programme n'a pas pu être chargé. Actualisez la page ou contactez-nous si le problème persiste.</p>`;
    }
    console.error(error);
  }
};

const initNavigation = () => {
  if (!navToggle || !navLinks) return;
  navToggle.addEventListener("click", () => {
    const isExpanded = navToggle.getAttribute("aria-expanded") === "true";
    navToggle.setAttribute("aria-expanded", String(!isExpanded));
    navLinks.classList.toggle("open");
  });

  navLinks.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", () => {
      navToggle.setAttribute("aria-expanded", "false");
      navLinks.classList.remove("open");
    });
  });
};

document.addEventListener("DOMContentLoaded", () => {
  createSnowflakes();
  createLanterns();
  initNavigation();
  loadItinerary();
});
