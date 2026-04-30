const STORAGE_KEY = "dnd-club-organizer-state";
// The Session Notes screen lets the organizer override this after deploying the backend.
let API_BASE_URL =
  localStorage.getItem("dnd-club-api-base-url") ||
  (["localhost", "127.0.0.1"].includes(window.location.hostname)
    ? "http://localhost:3001"
    : "https://dnd-club-session-organizer.onrender.com");
const MAX_AUDIO_UPLOAD_BYTES = 25 * 1024 * 1024;
const SUPPORTED_AUDIO_EXTENSIONS = [".mp3", ".mp4", ".mpeg", ".mpga", ".m4a", ".wav", ".webm"];

const initialState = {
  members: [],
  dms: [],
  campaigns: [],
  sessions: [],
  currentCampaignId: null,
  currentSessionId: null,
  settings: {
    manualMode: true,
    aiSetup: {
      lastCheckedAt: null,
      backendReachable: null,
      hasOpenAIKey: null,
      message: "AI setup has not been checked yet.",
      service: "",
    },
  },
  combat: {
    selectedMemberIds: [],
    monsters: [],
    combatants: [],
    activeIndex: 0,
    sorted: false,
    startedAt: null,
    actions: [],
    battleRunning: false,
    actionMode: "idle",
    selectedTargetId: null,
    selectedTargetIds: [],
    pendingDamageTargets: [],
    damageMode: "attack",
  },
};

let state = loadState();
let mediaRecorder = null;
let recordingChunks = [];
let recordingStartedAt = null;
let recordingTimerId = null;
let recordingDownloadUrl = null;
let audioUploadStatus = "ready";
let speakerTimingStatus = "ready";
let highlightedBattleId = null;

const elements = {
  views: document.querySelectorAll(".view"),
  navButtons: document.querySelectorAll("[data-view]"),
  homeTitle: document.querySelector("#home-title"),
  heroMemberCount: document.querySelector("#hero-member-count"),
  heroDmCount: document.querySelector("#hero-dm-count"),
  heroCampaignName: document.querySelector("#hero-campaign-name"),
  heroSessionCount: document.querySelector("#hero-session-count"),
  memberForm: document.querySelector("#member-form"),
  memberFormTitle: document.querySelector("#member-form-title"),
  memberId: document.querySelector("#member-id"),
  playerName: document.querySelector("#player-name"),
  characterName: document.querySelector("#character-name"),
  characterRace: document.querySelector("#character-race"),
  armorClass: document.querySelector("#armor-class"),
  maxHealth: document.querySelector("#max-health"),
  currentHealth: document.querySelector("#current-health"),
  cancelEdit: document.querySelector("#cancel-edit"),
  dmForm: document.querySelector("#dm-form"),
  dmFormTitle: document.querySelector("#dm-form-title"),
  dmId: document.querySelector("#dm-id"),
  dmName: document.querySelector("#dm-name"),
  cancelDmEdit: document.querySelector("#cancel-dm-edit"),
  rosterCount: document.querySelector("#roster-count"),
  rosterEmpty: document.querySelector("#roster-empty"),
  memberList: document.querySelector("#member-list"),
  dmCount: document.querySelector("#dm-count"),
  dmEmpty: document.querySelector("#dm-empty"),
  dmList: document.querySelector("#dm-list"),
  combatRosterEmpty: document.querySelector("#combat-roster-empty"),
  battleStripCampaign: document.querySelector("#battle-strip-campaign"),
  battleStripSession: document.querySelector("#battle-strip-session"),
  combatMemberOptions: document.querySelector("#combat-member-options"),
  monsterForm: document.querySelector("#monster-form"),
  monsterType: document.querySelector("#monster-type"),
  monsterCount: document.querySelector("#monster-count"),
  monsterArmorClass: document.querySelector("#monster-armor-class"),
  monsterHealth: document.querySelector("#monster-health"),
  monsterList: document.querySelector("#monster-list"),
  combatantsEmpty: document.querySelector("#combatants-empty"),
  combatantList: document.querySelector("#combatant-list"),
  autoRollInitiative: document.querySelector("#auto-roll-initiative"),
  sortInitiative: document.querySelector("#sort-initiative"),
  resetCombat: document.querySelector("#reset-combat"),
  currentTurn: document.querySelector("#current-turn"),
  turnDetail: document.querySelector("#turn-detail"),
  startInitiative: document.querySelector("#start-initiative"),
  battleResultCount: document.querySelector("#battle-result-count"),
  battleResultsEmpty: document.querySelector("#battle-results-empty"),
  battleResultsList: document.querySelector("#battle-results-list"),
  initiativeScreen: document.querySelector("#initiative-screen"),
  initiativeScreenTitle: document.querySelector("#initiative-screen-title"),
  initiativeScreenDetail: document.querySelector("#initiative-screen-detail"),
  initiativeScreenRoll: document.querySelector("#initiative-screen-roll"),
  initiativeScreenName: document.querySelector("#initiative-screen-name"),
  initiativeScreenType: document.querySelector("#initiative-screen-type"),
  initiativeScreenOrder: document.querySelector("#initiative-screen-order"),
  exitInitiativeScreen: document.querySelector("#exit-initiative-screen"),
  battleFocusStats: document.querySelector("#battle-focus-stats"),
  battleStatusList: document.querySelector("#battle-status-list"),
  battleActionMessage: document.querySelector("#battle-action-message"),
  battleActionButtons: document.querySelector("#battle-action-buttons"),
  battleAttack: document.querySelector("#battle-attack"),
  battleMultiAttack: document.querySelector("#battle-multi-attack"),
  battleOther: document.querySelector("#battle-other"),
  battleTargetList: document.querySelector("#battle-target-list"),
  battleResultButtons: document.querySelector("#battle-result-buttons"),
  battlePass: document.querySelector("#battle-pass"),
  battleFail: document.querySelector("#battle-fail"),
  battleSpell: document.querySelector("#battle-spell"),
  battleSpellSaveButtons: document.querySelector("#battle-spell-save-buttons"),
  spellSavePass: document.querySelector("#spell-save-pass"),
  spellSaveFail: document.querySelector("#spell-save-fail"),
  damageForm: document.querySelector("#damage-form"),
  damageAmount: document.querySelector("#damage-amount"),
  campaignForm: document.querySelector("#campaign-form"),
  campaignFormTitle: document.querySelector("#campaign-form-title"),
  campaignId: document.querySelector("#campaign-id"),
  campaignName: document.querySelector("#campaign-name"),
  cancelCampaignEdit: document.querySelector("#cancel-campaign-edit"),
  sessionForm: document.querySelector("#session-form"),
  activeCampaignNote: document.querySelector("#active-campaign-note"),
  sessionDate: document.querySelector("#session-date"),
  campaignCount: document.querySelector("#campaign-count"),
  campaignsEmpty: document.querySelector("#campaigns-empty"),
  campaignList: document.querySelector("#campaign-list"),
  sessionCount: document.querySelector("#session-count"),
  sessionsEmpty: document.querySelector("#sessions-empty"),
  sessionList: document.querySelector("#session-list"),
  speakingSessionContext: document.querySelector("#speaking-session-context"),
  recordingStatus: document.querySelector("#recording-status"),
  recordingDuration: document.querySelector("#recording-duration"),
  startRecording: document.querySelector("#start-recording"),
  stopRecording: document.querySelector("#stop-recording"),
  downloadRecording: document.querySelector("#download-recording"),
  recordingMessage: document.querySelector("#recording-message"),
  recordingNextSteps: document.querySelector("#recording-next-steps"),
  recordingPasteTranscript: document.querySelector("#recording-paste-transcript"),
  recordingOpenNotes: document.querySelector("#recording-open-notes"),
  speakingManualMode: document.querySelector("#speaking-manual-mode"),
  audioProcessingStatus: document.querySelector("#audio-processing-status"),
  audioProcessingNote: document.querySelector("#audio-processing-note"),
  audioProcessingFile: document.querySelector("#audio-processing-file"),
  speakerReviewStatus: document.querySelector("#speaker-review-status"),
  audioUploadFile: document.querySelector("#audio-upload-file"),
  transcribeAudio: document.querySelector("#transcribe-audio"),
  processSpeakerTiming: document.querySelector("#process-speaker-timing"),
  audioUploadMessage: document.querySelector("#audio-upload-message"),
  speakerCount: document.querySelector("#speaker-count"),
  speakerEmpty: document.querySelector("#speaker-empty"),
  speakerList: document.querySelector("#speaker-list"),
  saveSpeakerStats: document.querySelector("#save-speaker-stats"),
  speakerStatsMessage: document.querySelector("#speaker-stats-message"),
  speakingStripCampaign: document.querySelector("#speaking-strip-campaign"),
  speakingStripSession: document.querySelector("#speaking-strip-session"),
  manualTranscriptField: document.querySelector("#manual-transcript-field"),
  saveManualTranscript: document.querySelector("#save-manual-transcript"),
  manualTranscriptMessage: document.querySelector("#manual-transcript-message"),
  manualTranscriptOpenNotes: document.querySelector("#manual-transcript-open-notes"),
  sessionSpeakingChart: document.querySelector("#session-speaking-chart"),
  campaignSpeakingChart: document.querySelector("#campaign-speaking-chart"),
  speakerSegmentCount: document.querySelector("#speaker-segment-count"),
  speakerSegmentsEmpty: document.querySelector("#speaker-segments-empty"),
  speakerSegmentList: document.querySelector("#speaker-segment-list"),
  speakerMapList: document.querySelector("#speaker-map-list"),
  applySpeakerReview: document.querySelector("#apply-speaker-review"),
  speakerSegmentForm: document.querySelector("#speaker-segment-form"),
  segmentSpeakerLabel: document.querySelector("#segment-speaker-label"),
  segmentMinutes: document.querySelector("#segment-minutes"),
  segmentNote: document.querySelector("#segment-note"),
  progressSessionContext: document.querySelector("#progress-session-context"),
  progressStripCampaign: document.querySelector("#progress-strip-campaign"),
  progressStripSession: document.querySelector("#progress-strip-session"),
  progressForm: document.querySelector("#progress-form"),
  transcriptField: document.querySelector("#transcript-field"),
  recapSummaryField: document.querySelector("#recap-summary-field"),
  recapEventsField: document.querySelector("#recap-events-field"),
  recapNpcsField: document.querySelector("#recap-npcs-field"),
  recapLocationsField: document.querySelector("#recap-locations-field"),
  recapQuestsField: document.querySelector("#recap-quests-field"),
  recapThreadsField: document.querySelector("#recap-threads-field"),
  aiStatus: document.querySelector("#ai-status"),
  aiToolsNote: document.querySelector("#ai-tools-note"),
  generateRecap: document.querySelector("#generate-recap"),
  progressManualMode: document.querySelector("#progress-manual-mode"),
  apiBaseUrl: document.querySelector("#api-base-url"),
  saveApiBaseUrl: document.querySelector("#save-api-base-url"),
  clearApiBaseUrl: document.querySelector("#clear-api-base-url"),
  aiSetupStatus: document.querySelector("#ai-setup-status"),
  checkAiSetup: document.querySelector("#check-ai-setup"),
  aiSetupMessage: document.querySelector("#ai-setup-message"),
  aiSetupManualMode: document.querySelector("#ai-setup-manual-mode"),
  aiSetupBackendUrl: document.querySelector("#ai-setup-backend-url"),
  aiSetupBackendHealth: document.querySelector("#ai-setup-backend-health"),
  aiSetupOpenAiKey: document.querySelector("#ai-setup-openai-key"),
  aiSetupRecapReady: document.querySelector("#ai-setup-recap-ready"),
  aiSetupTranscriptionReady: document.querySelector("#ai-setup-transcription-ready"),
  aiSetupCheckedAt: document.querySelector("#ai-setup-checked-at"),
  storySessionCount: document.querySelector("#story-session-count"),
  exportSessionMarkdown: document.querySelector("#export-session-markdown"),
  exportCampaignMarkdown: document.querySelector("#export-campaign-markdown"),
  campaignStoryEmpty: document.querySelector("#campaign-story-empty"),
  campaignStoryList: document.querySelector("#campaign-story-list"),
};

elements.sessionDate.valueAsDate = new Date();

function loadState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    return saved ? mergeState(saved) : clone(initialState);
  } catch {
    return clone(initialState);
  }
}

function mergeState(saved) {
  const migrated = {
    ...clone(initialState),
    ...saved,
    settings: {
      ...clone(initialState.settings),
      ...(saved.settings || {}),
    },
    combat: {
      ...clone(initialState.combat),
      ...(saved.combat || {}),
    },
  };
  migrated.settings.aiSetup = {
    ...clone(initialState.settings.aiSetup),
    ...(saved.settings?.aiSetup || {}),
  };
  migrated.combat = normalizeCombat(migrated.combat);
  migrateCampaigns(migrated);
  return migrated;
}

function migrateCampaigns(nextState) {
  nextState.dms = Array.isArray(nextState.dms) ? nextState.dms : [];
  nextState.members = Array.isArray(nextState.members)
    ? nextState.members.map(normalizeMemberStats)
    : [];
  nextState.sessions = Array.isArray(nextState.sessions) ? nextState.sessions : [];
  nextState.campaigns = Array.isArray(nextState.campaigns) ? nextState.campaigns : [];

  nextState.sessions = nextState.sessions.map((session, index) => ({
    id: session.id || createId("session"),
    date: session.date || new Date().toISOString().slice(0, 10),
    campaignId: session.campaignId || null,
    participantMemberIds: session.participantMemberIds || [],
    combats: Array.isArray(session.combats) ? session.combats.map(normalizeBattleRecord) : [],
    recording: normalizeRecording(session.recording),
    transcript: session.transcript || "",
    speakerStats: Array.isArray(session.speakerStats) ? session.speakerStats : [],
    aiStatus: session.aiStatus || "not-ready",
    speakerSegments: Array.isArray(session.speakerSegments) ? session.speakerSegments : [],
    speakerReviewStatus: session.speakerReviewStatus || "manual",
    recap: normalizeRecap(session.recap),
    createdAt: session.createdAt || Date.now() + index,
  }));

  if (!nextState.campaigns.length && nextState.sessions.length) {
    const campaignId = createId("campaign");
    nextState.campaigns.push({
      id: campaignId,
      name: "First Campaign",
      dmIds: nextState.dms.map((dm) => dm.id),
      sessionIds: nextState.sessions.map((session) => session.id),
      status: "active",
    });
    nextState.sessions = nextState.sessions.map((session) => ({ ...session, campaignId }));
    nextState.currentCampaignId = campaignId;
  }

  nextState.campaigns = nextState.campaigns.map((campaign) => {
    const sessionIds = nextState.sessions
      .filter((session) => session.campaignId === campaign.id || campaign.sessionIds?.includes(session.id))
      .map((session) => session.id);
    return {
      id: campaign.id || createId("campaign"),
      name: campaign.name || "Untitled Campaign",
      dmIds: Array.isArray(campaign.dmIds) ? campaign.dmIds : [],
      sessionIds,
      status: campaign.status === "finished" ? "finished" : "active",
    };
  });

  const currentCampaign = nextState.campaigns.find((campaign) => campaign.id === nextState.currentCampaignId);
  if (!currentCampaign) {
    nextState.currentCampaignId =
      nextState.campaigns.find((campaign) => campaign.status === "active")?.id || nextState.campaigns[0]?.id || null;
  }

  const currentSession = nextState.sessions.find((session) => session.id === nextState.currentSessionId);
  if (!currentSession || currentSession.campaignId !== nextState.currentCampaignId) {
    nextState.currentSessionId = getCampaignSessions(nextState, nextState.currentCampaignId)[0]?.id || null;
  }
}

function normalizeCombat(combat) {
  return {
    ...clone(initialState.combat),
    ...combat,
    monsters: Array.isArray(combat.monsters) ? combat.monsters.map(normalizeCombatantStats) : [],
    combatants: Array.isArray(combat.combatants) ? combat.combatants.map(normalizeCombatantStats) : [],
    actions: Array.isArray(combat.actions) ? combat.actions : [],
    pendingDamageTargets: Array.isArray(combat.pendingDamageTargets) ? combat.pendingDamageTargets : [],
    selectedTargetIds: Array.isArray(combat.selectedTargetIds) ? combat.selectedTargetIds : [],
    damageMode: combat.damageMode || "attack",
  };
}

function normalizeCombatantStats(combatant) {
  const maxHealth = Number(combatant.maxHealth || combatant.health || 0);
  const currentHealth = combatant.currentHealth ?? maxHealth;
  const startingHealth = combatant.startingHealth ?? combatant.initialHealth ?? maxHealth;
  return {
    ...combatant,
    armorClass: combatant.armorClass ?? "",
    maxHealth: combatant.maxHealth ?? maxHealth,
    currentHealth: Math.max(0, Number(currentHealth || 0)),
    startingHealth: Math.max(0, Number(startingHealth || 0)),
    defeated: Boolean(combatant.defeated) || Number(currentHealth || 0) <= 0,
  };
}

function normalizeBattleRecord(record) {
  const combatants = Array.isArray(record.combatants) ? record.combatants.map(normalizeCombatantStats) : [];
  const finalCombatants = Array.isArray(record.finalCombatants)
    ? record.finalCombatants.map(normalizeCombatantStats)
    : combatants.map((combatant) => ({ ...combatant }));
  return {
    id: record.id || createId("battle"),
    name: record.name || "Battle",
    combatants,
    activeIndex: record.activeIndex || 0,
    status: record.status || "setup",
    startedAt: record.startedAt || null,
    endedAt: record.endedAt || null,
    actions: Array.isArray(record.actions) ? record.actions : [],
    finalCombatants,
    summary: record.summary || "",
  };
}

function normalizeMemberStats(member) {
  const maxHealth = member.maxHealth === "" || member.maxHealth === undefined ? "" : Number(member.maxHealth);
  const fallbackCurrent = maxHealth === "" ? "" : maxHealth;
  const currentHealth =
    member.currentHealth === "" || member.currentHealth === undefined ? fallbackCurrent : Number(member.currentHealth);
  return {
    ...member,
    armorClass: member.armorClass ?? "",
    maxHealth,
    currentHealth: maxHealth === "" ? "" : clampHealth(currentHealth, maxHealth),
  };
}

function clampHealth(currentHealth, maxHealth) {
  const max = Math.max(0, Number(maxHealth || 0));
  return Math.min(max, Math.max(0, Number(currentHealth || 0)));
}

function clone(value) {
  if (typeof structuredClone === "function") {
    return structuredClone(value);
  }
  return JSON.parse(JSON.stringify(value));
}

function normalizeRecording(recording) {
  if (!recording) return null;
  return {
    status: recording.status || "saved",
    startedAt: recording.startedAt || null,
    durationSeconds: Number(recording.durationSeconds || 0),
    fileName: recording.fileName || "",
  };
}

function normalizeRecap(recap) {
  if (typeof recap === "string") {
    return {
      summary: recap,
      events: "",
      npcs: "",
      locations: "",
      quests: "",
      unresolvedThreads: "",
    };
  }
  return {
    summary: recap?.summary || "",
    events: recap?.events || "",
    npcs: recap?.npcs || "",
    locations: recap?.locations || "",
    quests: recap?.quests || "",
    unresolvedThreads: recap?.unresolvedThreads || "",
  };
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function createId(prefix) {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function showView(viewName) {
  elements.views.forEach((view) => {
    view.classList.toggle("active", view.id === `view-${viewName}`);
  });
  document.querySelectorAll(".nav-tab").forEach((button) => {
    button.classList.toggle("active", button.dataset.view === viewName);
  });
}

function render() {
  renderHome();
  renderSessionStrips();
  renderRoster();
  renderDms();
  renderCombatBuilder();
  renderCombatants();
  renderBattleResults();
  renderCampaignsAndSessions();
  renderSpeakingTime();
  renderProgressNotes();
  saveState();
}

function renderSessionStrips() {
  const campaign = getActiveCampaign();
  const session = getCurrentSession();
  const sessionText = session ? getSessionLabel(session, getSessionIndex(session)) : "No session selected";
  const campaignText = campaign?.name || "No active campaign";
  [
    [elements.battleStripCampaign, elements.battleStripSession],
    [elements.speakingStripCampaign, elements.speakingStripSession],
    [elements.progressStripCampaign, elements.progressStripSession],
  ].forEach(([campaignElement, sessionElement]) => {
    if (campaignElement) campaignElement.textContent = campaignText;
    if (sessionElement) sessionElement.textContent = sessionText;
  });
}

function renderHome() {
  const activeCampaign = getActiveCampaign();
  const activeSessions = getCampaignSessions(state, state.currentCampaignId);
  elements.homeTitle.textContent = getCampaignDisplayTitle(activeCampaign);
  elements.heroMemberCount.textContent = state.members.length;
  elements.heroDmCount.textContent = state.dms.length;
  elements.heroCampaignName.textContent = activeCampaign?.name || "No campaign";
  elements.heroSessionCount.textContent = activeSessions.length;
}

function renderRoster() {
  elements.rosterCount.textContent = `${state.members.length} saved`;
  elements.rosterEmpty.classList.toggle("hidden", state.members.length > 0);

  elements.memberList.innerHTML = state.members
    .map(
      (member) => `
        <article class="member-card">
          <div>
            <strong>${escapeHtml(member.characterName)}</strong>
            <small>${escapeHtml(member.playerName)} playing a ${escapeHtml(member.characterRace)} - AC ${escapeHtml(member.armorClass || "unset")} - HP ${escapeHtml(member.currentHealth ?? "unset")}/${escapeHtml(member.maxHealth || "unset")}</small>
          </div>
          <div class="card-actions">
            <button class="icon-button" type="button" data-edit-member="${member.id}" aria-label="Edit ${escapeHtml(member.characterName)}">Edit</button>
            <button class="icon-button delete-button" type="button" data-delete-member="${member.id}" aria-label="Delete ${escapeHtml(member.characterName)}">Del</button>
          </div>
        </article>
      `,
    )
    .join("");
}

function renderDms() {
  elements.dmCount.textContent = `${state.dms.length} saved`;
  elements.dmEmpty.classList.toggle("hidden", state.dms.length > 0);
  elements.dmList.innerHTML = state.dms
    .map(
      (dm) => `
        <article class="member-card">
          <div>
            <strong>${escapeHtml(dm.name)}</strong>
            <small>Dungeon Master</small>
          </div>
          <div class="card-actions">
            <button class="icon-button" type="button" data-edit-dm="${dm.id}" aria-label="Edit ${escapeHtml(dm.name)}">Edit</button>
            <button class="icon-button delete-button" type="button" data-delete-dm="${dm.id}" aria-label="Delete ${escapeHtml(dm.name)}">Del</button>
          </div>
        </article>
      `,
    )
    .join("");
}

function renderCombatBuilder() {
  elements.combatRosterEmpty.classList.toggle("hidden", state.members.length > 0);
  elements.combatMemberOptions.innerHTML = state.members
    .map((member) => {
      const checked = state.combat.selectedMemberIds.includes(member.id) ? "checked" : "";
      const ready = Number(member.armorClass) > 0 && Number(member.maxHealth) > 0 && member.currentHealth !== "";
      return `
        <label class="check-row ${ready ? "" : "disabled-row"}">
          <input type="checkbox" data-combat-member="${member.id}" ${checked} ${ready ? "" : "disabled"} />
          <span>
            <strong>${escapeHtml(member.characterName)}</strong><br />
            <span class="combatant-meta">${escapeHtml(member.playerName)} - ${escapeHtml(member.characterRace)} - AC ${escapeHtml(member.armorClass || "unset")} - HP ${escapeHtml(member.currentHealth ?? "unset")}/${escapeHtml(member.maxHealth || "unset")}${ready ? "" : " - edit roster stats before battle"}</span>
          </span>
        </label>
      `;
    })
    .join("");

  elements.monsterList.innerHTML = state.combat.monsters
    .map(
      (monster) => `
        <div class="monster-chip">
          <span>${escapeHtml(monster.displayName)} - AC ${escapeHtml(monster.armorClass)} - HP ${escapeHtml(monster.maxHealth)}</span>
          <button class="icon-button delete-button" data-delete-monster="${monster.id}" type="button">Del</button>
        </div>
      `,
    )
    .join("");
}

function renderCombatants() {
  syncCombatants();
  const combatants = state.combat.combatants;
  elements.combatantsEmpty.classList.toggle("hidden", combatants.length > 0);

  elements.combatantList.innerHTML = combatants
    .map((combatant, index) => {
      const active = state.combat.sorted && index === state.combat.activeIndex ? "active" : "";
      const tied = hasTie(combatant) ? "tie" : "";
      return `
        <div class="combatant-row ${active} ${tied}">
          <span class="turn-number">${index + 1}</span>
          <div>
            <strong>${escapeHtml(combatant.displayName)}</strong>
            <div class="combatant-meta">${combatant.type === "character" ? "Character" : "Monster"} - AC ${escapeHtml(combatant.armorClass || "unset")} - HP ${escapeHtml(combatant.currentHealth ?? combatant.maxHealth ?? "unset")}/${escapeHtml(combatant.maxHealth || "unset")}${tied ? " - tied roll" : ""}</div>
          </div>
          <label>
            Roll
            <input type="number" data-roll="${combatant.id}" value="${combatant.initiativeRoll ?? ""}" />
          </label>
          <div class="reorder-buttons" aria-label="Manual tie order controls">
            <button class="icon-button" data-move-up="${combatant.id}" type="button" aria-label="Move up">Up</button>
            <button class="icon-button" data-move-down="${combatant.id}" type="button" aria-label="Move down">Down</button>
          </div>
        </div>
      `;
    })
    .join("");

  renderCurrentTurn();
  renderInitiativeScreen();
  saveCombatToSession();
}

function renderBattleResults() {
  const battles = getCompletedBattleResults();
  elements.battleResultCount.textContent = `${battles.length} saved`;
  elements.battleResultsEmpty.classList.toggle("hidden", battles.length > 0);
  elements.battleResultsList.innerHTML = battles
    .map(({ battle, session, sessionIndex }) => {
      const defeated = (battle.finalCombatants || []).filter(
        (combatant) => combatant.defeated || Number(combatant.currentHealth) <= 0,
      );
      const actions = (battle.actions || []).slice(-4);
      return `
        <article class="battle-result-card ${battle.id === highlightedBattleId ? "highlighted" : ""}" id="battle-result-${battle.id}">
          <div class="battle-result-header">
            <div>
              <strong>${escapeHtml(battle.name || "Saved battle")}</strong>
              <small>${escapeHtml(getSessionLabel(session, sessionIndex))} - ${escapeHtml(formatDate((battle.endedAt || battle.startedAt || session.date).slice(0, 10)))}</small>
            </div>
            <div class="button-row compact">
              <button class="icon-button" type="button" data-copy-battle="${battle.id}" data-battle-session="${session.id}">Copy to notes</button>
              <button class="icon-button delete-button" type="button" data-delete-battle="${battle.id}" data-battle-session="${session.id}">Del</button>
            </div>
          </div>
          <div class="battle-result-meta">
            <span>${(battle.finalCombatants || []).length} combatants</span>
            <span>${defeated.length} defeated</span>
          </div>
          <div class="battle-result-combatants">
            ${(battle.finalCombatants || [])
              .map((combatant) => {
                const starting = combatant.startingHealth ?? combatant.maxHealth;
                return `
                  <span>
                    <strong>${escapeHtml(combatant.displayName)}</strong>
                    ${escapeHtml(combatant.type === "character" ? "Character" : "Monster")} - HP ${escapeHtml(starting)}/${escapeHtml(combatant.maxHealth)} to ${escapeHtml(combatant.currentHealth)}/${escapeHtml(combatant.maxHealth)}
                  </span>
                `;
              })
              .join("")}
          </div>
          <div class="battle-result-actions">
            ${
              actions.length
                ? actions.map((action) => `<span>${escapeHtml(action.text)}</span>`).join("")
                : "<span>No action notes saved.</span>"
            }
          </div>
          <pre class="battle-markdown-preview">${escapeHtml(buildBattleSummaryMarkdown(battle))}</pre>
        </article>
      `;
    })
    .join("");
}

function renderCurrentTurn() {
  const combatants = state.combat.combatants;
  if (!combatants.length) {
    elements.currentTurn.textContent = "No active battle";
    elements.turnDetail.textContent = "Select characters or add monsters to begin.";
    elements.startInitiative.disabled = true;
    return;
  }

  if (!state.combat.sorted) {
    elements.currentTurn.textContent = "Ready to sort";
    elements.turnDetail.textContent = "Enter rolls, then sort or auto-roll before starting.";
    elements.startInitiative.disabled = true;
    return;
  }

  if (!hasCompleteInitiative()) {
    elements.currentTurn.textContent = "Enter every roll";
    elements.turnDetail.textContent = "Every combatant needs a turn-order roll before battle can start.";
    elements.startInitiative.disabled = true;
    return;
  }

  ensureActiveLivingCombatant();
  const active = combatants[state.combat.activeIndex] || combatants[0];
  if (!active || !getLivingCombatants().length) {
    elements.currentTurn.textContent = "Battle complete";
    elements.turnDetail.textContent = "All combatants are at 0 health.";
    elements.startInitiative.disabled = true;
    return;
  }
  elements.currentTurn.textContent = "Start Battle";
  elements.turnDetail.textContent = `${active.displayName} is first with roll ${active.initiativeRoll ?? "not entered"}.`;
  elements.startInitiative.disabled = false;
}

function renderInitiativeScreen() {
  const combatants = state.combat.combatants;
  ensureActiveLivingCombatant();
  const active = getActiveCombatant();
  if (!combatants.length || !state.combat.sorted || !active) {
    elements.initiativeScreenTitle.textContent = "No active turn";
    elements.initiativeScreenDetail.textContent = "Start battle from the tracker.";
    elements.initiativeScreenRoll.textContent = "Roll --";
    elements.initiativeScreenName.textContent = "No combatant";
    elements.initiativeScreenType.textContent = "Waiting for battle";
    elements.battleFocusStats.innerHTML = "";
    elements.initiativeScreenOrder.innerHTML = "";
    elements.battleStatusList.innerHTML = "";
    resetBattleActionUi("Choose an action for the active combatant.");
    return;
  }

  elements.initiativeScreenTitle.textContent = `${active.displayName}'s turn`;
  elements.initiativeScreenDetail.textContent = `Turn ${state.combat.activeIndex + 1} of ${combatants.length}`;
  elements.initiativeScreenRoll.textContent = `Roll ${active.initiativeRoll ?? "--"}`;
  elements.initiativeScreenName.textContent = active.displayName;
  elements.initiativeScreenType.textContent = active.type === "character" ? "Player character" : "Monster";
  elements.battleFocusStats.innerHTML = `
    <span>AC ${escapeHtml(active.armorClass)}</span>
    <span>HP ${escapeHtml(active.currentHealth)}/${escapeHtml(active.maxHealth)}</span>
  `;
  elements.initiativeScreenOrder.innerHTML = combatants
    .map(
      (combatant, index) => `
        <div class="screen-order-row ${index === state.combat.activeIndex ? "active" : ""} ${combatant.defeated || combatant.currentHealth <= 0 ? "defeated" : ""}">
          <span>${index + 1}</span>
          <strong>${escapeHtml(combatant.displayName)}</strong>
          <small>${combatant.initiativeRoll ?? "--"} / HP ${escapeHtml(combatant.currentHealth)}</small>
        </div>
      `,
    )
    .join("");
  elements.battleStatusList.innerHTML = combatants
    .map(
      (combatant) => `
        <div class="battle-status-row ${combatant.defeated || combatant.currentHealth <= 0 ? "defeated" : ""}">
          <strong>${escapeHtml(combatant.displayName)}</strong>
          <span>AC ${escapeHtml(combatant.armorClass)}</span>
          <span>HP ${escapeHtml(combatant.currentHealth)}/${escapeHtml(combatant.maxHealth)}</span>
        </div>
      `,
    )
    .join("");
  renderBattleActionUi();
}

function renderCampaignsAndSessions() {
  const activeCampaign = getActiveCampaign();
  const activeSessions = getCampaignSessions(state, state.currentCampaignId);
  elements.campaignCount.textContent = `${state.campaigns.length} saved`;
  elements.campaignsEmpty.classList.toggle("hidden", state.campaigns.length > 0);
  elements.activeCampaignNote.textContent = activeCampaign
    ? `New sessions will be saved under ${activeCampaign.name}.`
    : "Create a campaign first.";
  elements.sessionForm.querySelector("button").disabled = !activeCampaign;

  elements.campaignList.innerHTML = state.campaigns
    .map((campaign) => {
      const campaignSessions = getCampaignSessions(state, campaign.id);
      const dms = getCampaignDms(campaign);
      const current = campaign.id === state.currentCampaignId ? "Active" : "Set active";
      const finished = campaign.status === "finished";
      return `
        <article class="campaign-card">
          <div class="campaign-main">
            <strong>${escapeHtml(campaign.name)}</strong>
            <div class="campaign-meta">
              <span>${finished ? "Finished" : "Active"}</span>
              <span>${campaignSessions.length} session${campaignSessions.length === 1 ? "" : "s"}</span>
              <span>DM: ${dms.length ? escapeHtml(formatNameList(dms.map((dm) => dm.name))) : "None yet"}</span>
            </div>
          </div>
          <div class="campaign-actions">
            <button class="icon-button" type="button" data-current-campaign="${campaign.id}">${current}</button>
            <button class="icon-button" type="button" data-finish-campaign="${campaign.id}">${finished ? "Reopen" : "Finish"}</button>
            <button class="icon-button" type="button" data-edit-campaign="${campaign.id}" aria-label="Edit ${escapeHtml(campaign.name)}">Edit</button>
            <button class="icon-button delete-button" type="button" data-delete-campaign="${campaign.id}" aria-label="Delete ${escapeHtml(campaign.name)}">Del</button>
          </div>
        </article>
      `;
    })
    .join("");

  elements.sessionCount.textContent = `${activeSessions.length} saved`;
  elements.sessionsEmpty.classList.toggle("hidden", activeSessions.length > 0);
  elements.sessionList.innerHTML = activeSessions
    .map((session, index) => {
      const current = session.id === state.currentSessionId ? "Current session" : "Make current";
      const completedBattles = (session.combats || []).filter((battle) => battle.id !== "current-combat" && battle.status === "completed");
      const combatCount = completedBattles.length;
      return `
        <article class="session-card">
          <div>
            <strong>${escapeHtml(getSessionLabel(session, index))}</strong>
            <small>${combatCount} battle record${combatCount === 1 ? "" : "s"}</small>
            ${
              combatCount
                ? `<div class="battle-summary-list">${completedBattles
                    .map(
                      (battle) => `
                        <span>
                          ${escapeHtml(battle.summary || battle.name || "Saved battle")}
                          <button class="inline-link-button" type="button" data-view-battle="${battle.id}" data-battle-session="${session.id}">View results</button>
                        </span>
                      `,
                    )
                    .join("")}</div>`
                : ""
            }
          </div>
          <div class="card-actions">
            <button class="icon-button" type="button" data-current-session="${session.id}">${current}</button>
            <button class="icon-button delete-button" type="button" data-delete-session="${session.id}">Del</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderSpeakingTime() {
  const session = getCurrentSession();
  const activeCampaign = getActiveCampaign();
  const speakers = getSpeakers();
  const manualMode = Boolean(state.settings.manualMode);
  const hasAudioFile = Boolean(elements.audioUploadFile.files?.[0]);
  elements.speakingSessionContext.textContent = session
    ? `Speaking time for ${getSessionLabel(session, getSessionIndex(session))} in ${activeCampaign?.name || "the active campaign"}.`
    : "Create or select a session to record audio and track speaking time.";

  updateManualModeButton(elements.speakingManualMode);
  const recording = session?.recording || null;
  elements.recordingStatus.textContent = recording?.status || "No recording";
  elements.recordingDuration.textContent = formatDuration(recording?.durationSeconds || 0);
  elements.audioProcessingStatus.textContent = getAudioProcessingStatus(session);
  elements.audioProcessingFile.textContent = recording?.fileName || "No file yet";
  elements.speakerReviewStatus.textContent = formatStatus(session?.speakerReviewStatus || "manual");
  elements.audioProcessingNote.textContent = getAudioProcessingMessage(session);
  elements.startRecording.disabled = !session || Boolean(mediaRecorder);
  elements.stopRecording.disabled = !mediaRecorder;
  elements.recordingNextSteps.classList.toggle("hidden", recording?.status !== "ready to download");
  elements.audioUploadFile.disabled =
    !session ||
    manualMode ||
    audioUploadStatus === "uploading" ||
    audioUploadStatus === "transcribing" ||
    speakerTimingStatus === "processing";
  elements.transcribeAudio.disabled =
    !session ||
    manualMode ||
    !hasAudioFile ||
    audioUploadStatus === "uploading" ||
    audioUploadStatus === "transcribing" ||
    speakerTimingStatus === "processing";
  elements.transcribeAudio.textContent =
    !session
      ? "Select session first"
      : manualMode
      ? "Manual Mode enabled"
      : !hasAudioFile
      ? "Choose file first"
      : audioUploadStatus === "uploading" || audioUploadStatus === "transcribing"
      ? "Transcribing..."
      : "Upload and transcribe";
  elements.processSpeakerTiming.disabled =
    !session ||
    manualMode ||
    !hasAudioFile ||
    audioUploadStatus === "uploading" ||
    audioUploadStatus === "transcribing" ||
    speakerTimingStatus === "processing";
  elements.processSpeakerTiming.textContent =
    !session
      ? "Select session first"
      : manualMode
      ? "Manual Mode enabled"
      : !hasAudioFile
      ? "Choose file first"
      : speakerTimingStatus === "processing"
      ? "Processing speakers..."
      : "Process speaker timing";
  elements.manualTranscriptField.disabled = !session;
  elements.manualTranscriptField.value = session?.transcript || "";
  elements.saveManualTranscript.disabled = !session;
  elements.manualTranscriptOpenNotes.disabled = !session;
  elements.recordingPasteTranscript.disabled = !session;
  elements.recordingOpenNotes.disabled = !session;
  elements.saveSpeakerStats.disabled = !session;
  elements.speakerCount.textContent = `${speakers.length} speaker${speakers.length === 1 ? "" : "s"}`;
  elements.speakerEmpty.classList.toggle("hidden", speakers.length > 0);

  if (!recordingDownloadUrl) {
    elements.downloadRecording.classList.add("hidden");
    elements.downloadRecording.removeAttribute("href");
  }

  const stats = session?.speakerStats || [];
  elements.speakerList.innerHTML = speakers
    .map((speaker) => {
      const stat = stats.find((entry) => entry.speakerId === speaker.id && entry.speakerType === speaker.type);
      return `
        <label class="speaker-row">
          <span>
            <strong>${escapeHtml(speaker.name)}</strong>
            <small>${speaker.type === "dm" ? "Dungeon Master" : "Player character"}</small>
          </span>
          <input type="number" min="0" step="1" data-speaker-id="${speaker.id}" data-speaker-type="${speaker.type}" value="${stat?.minutes ?? 0}" ${session ? "" : "disabled"} />
        </label>
      `;
    })
    .join("");

  renderSpeakingCharts(session, activeCampaign);
  renderSpeakerSegments(session);
}

function renderProgressNotes() {
  const session = getCurrentSession();
  const activeCampaign = getActiveCampaign();
  const activeSessions = getCampaignSessions(state, state.currentCampaignId);
  const manualMode = Boolean(state.settings.manualMode);
  elements.progressSessionContext.textContent = session
    ? `Notes for ${getSessionLabel(session, getSessionIndex(session))} in ${activeCampaign?.name || "the active campaign"}.`
    : "Create or select a session to write transcripts, recaps, and campaign notes.";
  updateManualModeButton(elements.progressManualMode);
  elements.apiBaseUrl.value = API_BASE_URL;
  elements.aiStatus.textContent = formatStatus(session?.aiStatus || "not-ready");
  elements.aiToolsNote.textContent = session
    ? getAiToolsMessage(session)
    : "Create or select a session before preparing AI recap generation.";

  const recap = normalizeRecap(session?.recap);
  elements.transcriptField.value = session?.transcript || "";
  elements.recapSummaryField.value = recap.summary;
  elements.recapEventsField.value = recap.events;
  elements.recapNpcsField.value = recap.npcs;
  elements.recapLocationsField.value = recap.locations;
  elements.recapQuestsField.value = recap.quests;
  elements.recapThreadsField.value = recap.unresolvedThreads;

  const disabled = !session;
  [
    elements.transcriptField,
    elements.recapSummaryField,
    elements.recapEventsField,
    elements.recapNpcsField,
    elements.recapLocationsField,
    elements.recapQuestsField,
    elements.recapThreadsField,
  ].forEach((field) => {
    field.disabled = disabled;
  });
  elements.progressForm.querySelector("button").disabled = disabled;
  elements.exportSessionMarkdown.disabled = !session;
  elements.exportCampaignMarkdown.disabled = !activeCampaign;
  updateGenerateRecapButton();
  renderAiSetupStatus(session);

  const sessionsWithRecaps = activeSessions.filter((entry) => normalizeRecap(entry.recap).summary.trim());
  elements.storySessionCount.textContent = `${sessionsWithRecaps.length} session${sessionsWithRecaps.length === 1 ? "" : "s"}`;
  elements.campaignStoryEmpty.classList.toggle("hidden", sessionsWithRecaps.length > 0);
  elements.campaignStoryList.innerHTML = sessionsWithRecaps
    .map((entry) => {
      const entryRecap = normalizeRecap(entry.recap);
      return `
        <article class="story-card">
          <strong>${escapeHtml(getSessionLabel(entry, getSessionIndex(entry)))}</strong>
          <p>${escapeHtml(entryRecap.summary)}</p>
        </article>
      `;
    })
    .join("");
}

function renderSpeakingCharts(session, activeCampaign) {
  const sessionStats = session?.speakerStats || [];
  renderChart(elements.sessionSpeakingChart, sessionStats, "No speaking minutes saved for this session.", "session");

  const campaignTotals = new Map();
  getCampaignSessions(state, activeCampaign?.id).forEach((entry) => {
    (entry.speakerStats || []).forEach((stat) => {
      const key = `${stat.speakerType}:${stat.speakerId}`;
      const existing = campaignTotals.get(key) || { ...stat, minutes: 0 };
      existing.minutes += Number(stat.minutes || 0);
      campaignTotals.set(key, existing);
    });
  });
  renderChart(
    elements.campaignSpeakingChart,
    Array.from(campaignTotals.values()),
    "No speaking minutes saved for this campaign.",
    "campaign",
  );
}

function renderChart(container, stats, emptyText, scope) {
  const visibleStats = stats.filter((stat) => Number(stat.minutes) > 0);
  if (!visibleStats.length) {
    container.innerHTML = `<div class="empty-state">${emptyText}</div>`;
    return;
  }
  const maxMinutes = Math.max(...visibleStats.map((stat) => Number(stat.minutes || 0)), 1);
  container.innerHTML = visibleStats
    .map((stat) => {
      const width = Math.max(6, Math.round((Number(stat.minutes || 0) / maxMinutes) * 100));
      return `
        <div class="chart-row">
          <div class="chart-label">
            <strong>${escapeHtml(stat.name)}</strong>
            <span>${escapeHtml(stat.speakerType === "dm" ? "DM" : "Player")}</span>
          </div>
          <div class="chart-track"><span style="width: ${width}%"></span></div>
          <strong>${escapeHtml(stat.minutes)} min</strong>
          <button class="icon-button delete-button" type="button" data-delete-speaking-stat="${scope}" data-speaker-id="${escapeHtml(stat.speakerId)}" data-speaker-type="${escapeHtml(stat.speakerType)}">Del</button>
        </div>
      `;
    })
    .join("");
}

function renderSpeakerSegments(session) {
  const segments = session?.speakerSegments || [];
  const speakers = getSpeakers();
  elements.speakerSegmentForm.querySelector("button").disabled = !session;
  elements.applySpeakerReview.disabled = !session || !segments.length;
  [elements.segmentSpeakerLabel, elements.segmentMinutes, elements.segmentNote].forEach((field) => {
    field.disabled = !session;
  });
  elements.speakerSegmentCount.textContent = `${segments.length} segment${segments.length === 1 ? "" : "s"}`;
  elements.speakerSegmentsEmpty.classList.toggle("hidden", segments.length > 0);
  const aiLabels = Array.from(new Set(segments.filter((segment) => segment.source === "ai").map((segment) => segment.speakerLabel)));
  elements.speakerMapList.innerHTML = aiLabels.length
    ? aiLabels
        .map((label) => {
          const assignedSegment = segments.find((segment) => segment.speakerLabel === label && segment.assignedSpeakerId);
          const currentValue = assignedSegment
            ? `${assignedSegment.assignedSpeakerType}:${assignedSegment.assignedSpeakerId}`
            : "";
          return `
            <label class="speaker-map-row">
              <span>
                <strong>${escapeHtml(label || "Unknown speaker")}</strong>
                <small>${segments.filter((segment) => segment.speakerLabel === label).length} segment${segments.filter((segment) => segment.speakerLabel === label).length === 1 ? "" : "s"}</small>
              </span>
              <select data-speaker-map-label="${escapeHtml(label || "Unknown speaker")}">
                <option value="">Choose player or DM</option>
                ${speakers
                  .map(
                    (speaker) =>
                      `<option value="${speaker.type}:${speaker.id}" ${currentValue === `${speaker.type}:${speaker.id}` ? "selected" : ""}>${escapeHtml(speaker.name)} (${speaker.type === "dm" ? "DM" : "Player"})</option>`,
                  )
                  .join("")}
              </select>
            </label>
          `;
        })
        .join("")
    : "";
  elements.speakerSegmentList.innerHTML = segments
    .map(
      (segment) => `
        <article class="speaker-segment-card">
          <div>
            <strong>${escapeHtml(segment.assignedSpeakerName || segment.speakerLabel || "Unknown speaker")}</strong>
            <small>${escapeHtml(segment.source === "ai" ? `${segment.speakerLabel || "AI speaker"} · ${formatMinutes(segment.minutes || 0)} min` : `${formatMinutes(segment.minutes || 0)} min`)}</small>
          </div>
          <p>${escapeHtml(segment.note || segment.text || "Manual review segment")}</p>
          <button class="icon-button delete-button" type="button" data-delete-segment="${segment.id}">Del</button>
        </article>
      `,
    )
    .join("");
}

function syncCombatants() {
  const characterCombatants = state.combat.selectedMemberIds
    .map((memberId) => {
      const member = state.members.find((entry) => entry.id === memberId);
      if (!member) return null;
      const existing = state.combat.combatants.find((entry) => entry.memberId === memberId);
      return {
        id: existing?.id || `character-${memberId}`,
        type: "character",
        displayName: member.characterName,
        memberId,
        armorClass: Number(member.armorClass || existing?.armorClass || 0),
        maxHealth: Number(member.maxHealth || existing?.maxHealth || 0),
        currentHealth: existing?.currentHealth ?? Number(member.currentHealth ?? member.maxHealth ?? 0),
        startingHealth: existing?.startingHealth ?? Number(member.currentHealth ?? member.maxHealth ?? 0),
        defeated: Boolean(existing?.defeated) || Number(existing?.currentHealth ?? member.currentHealth ?? member.maxHealth ?? 0) <= 0,
        initiativeRoll: existing?.initiativeRoll ?? "",
      };
    })
    .filter(Boolean);

  const monsterCombatants = state.combat.monsters.map((monster) => {
    const existing = state.combat.combatants.find((entry) => entry.id === monster.id);
    return {
      id: monster.id,
      type: "monster",
      displayName: monster.displayName,
      armorClass: Number(monster.armorClass || existing?.armorClass || 0),
      maxHealth: Number(monster.maxHealth || existing?.maxHealth || 0),
      currentHealth: existing?.currentHealth ?? Number(monster.maxHealth || 0),
      startingHealth: existing?.startingHealth ?? Number(monster.maxHealth || 0),
      defeated: Boolean(existing?.defeated) || Number(existing?.currentHealth ?? monster.maxHealth ?? 0) <= 0,
      initiativeRoll: existing?.initiativeRoll ?? "",
    };
  });

  const nextCombatants = [...characterCombatants, ...monsterCombatants];
  if (state.combat.sorted) {
    const order = new Map(state.combat.combatants.map((entry, index) => [entry.id, index]));
    nextCombatants.sort((a, b) => (order.get(a.id) ?? 999) - (order.get(b.id) ?? 999));
  }
  state.combat.combatants = nextCombatants;
  if (state.combat.activeIndex >= nextCombatants.length) {
    state.combat.activeIndex = 0;
  }
}

function hasTie(combatant) {
  if (combatant.initiativeRoll === "" || combatant.initiativeRoll === null) return false;
  return state.combat.combatants.some(
    (entry) => entry.id !== combatant.id && Number(entry.initiativeRoll) === Number(combatant.initiativeRoll),
  );
}

function hasCompleteInitiative() {
  return (
    state.combat.combatants.length > 0 &&
    state.combat.combatants.every((combatant) => combatant.initiativeRoll !== "" && combatant.initiativeRoll !== null)
  );
}

function getActiveCombatant() {
  return state.combat.combatants[state.combat.activeIndex] || null;
}

function getLivingCombatants() {
  return state.combat.combatants.filter((combatant) => !combatant.defeated && Number(combatant.currentHealth) > 0);
}

function ensureActiveLivingCombatant() {
  const combatants = state.combat.combatants;
  if (!combatants.length || !getLivingCombatants().length) return;
  const active = combatants[state.combat.activeIndex];
  if (active && !active.defeated && Number(active.currentHealth) > 0) return;
  advanceTurn(false);
}

function advanceTurn(shouldRender = true) {
  const combatants = state.combat.combatants;
  if (!combatants.length) return;
  if (!getLivingCombatants().length) {
    resetBattleActionState();
    if (shouldRender) render();
    return;
  }
  for (let step = 1; step <= combatants.length; step += 1) {
    const nextIndex = (state.combat.activeIndex + step) % combatants.length;
    const next = combatants[nextIndex];
    if (next && !next.defeated && Number(next.currentHealth) > 0) {
      state.combat.activeIndex = nextIndex;
      break;
    }
  }
  resetBattleActionState();
  if (shouldRender) render();
}

function resetBattleActionState() {
  state.combat.actionMode = "idle";
  state.combat.selectedTargetId = null;
  state.combat.selectedTargetIds = [];
  state.combat.pendingDamageTargets = [];
  state.combat.damageMode = "attack";
}

function getAvailableTargets() {
  const active = getActiveCombatant();
  if (!active) return [];
  const targetType = active.type === "character" ? "monster" : "character";
  return state.combat.combatants.filter(
    (combatant) => combatant.type === targetType && !combatant.defeated && Number(combatant.currentHealth) > 0,
  );
}

function renderBattleActionUi() {
  const mode = state.combat.actionMode || "idle";
  elements.battleActionButtons.classList.toggle("hidden", !["idle"].includes(mode));
  elements.battleResultButtons.classList.toggle("hidden", mode !== "target-selected");
  elements.battleSpellSaveButtons.classList.toggle("hidden", mode !== "spell-save");
  elements.damageForm.classList.toggle("hidden", mode !== "damage");

  if (mode === "idle") {
    resetBattleActionUi("Choose an action for the active combatant.");
    return;
  }

  if (mode === "attack-target" || mode === "multi-target") {
    const targets = getAvailableTargets();
    const multi = mode === "multi-target";
    elements.battleActionMessage.textContent = multi
      ? "Choose one or more targets for the multi-attack."
      : "Choose one target to attack.";
    elements.battleTargetList.innerHTML = targets
      .map(
        (target) => `
          <label class="battle-target-row">
            <input type="${multi ? "checkbox" : "radio"}" name="battle-target" data-battle-target="${target.id}" ${state.combat.selectedTargetIds.includes(target.id) || state.combat.selectedTargetId === target.id ? "checked" : ""} />
            <span>
              <strong>${escapeHtml(target.displayName)}</strong>
              <small>AC ${escapeHtml(target.armorClass)} - HP ${escapeHtml(target.currentHealth)}/${escapeHtml(target.maxHealth)}</small>
            </span>
          </label>
        `,
      )
      .join("");
    if (multi) {
      elements.battleTargetList.insertAdjacentHTML("beforeend", '<button class="primary-button" type="button" id="confirm-multi-targets">Confirm targets</button>');
    }
    return;
  }

  if (mode === "target-selected") {
    const targets = getSelectedBattleTargets();
    elements.battleActionMessage.textContent = targets.length
      ? `${targets.map((target) => `${target.displayName}'s Armor Class is ${target.armorClass}`).join(". ")}. Did the roll pass, fail, or use a spell?`
      : "Choose the roll result.";
    elements.battleTargetList.innerHTML = "";
    return;
  }

  if (mode === "spell-save") {
    const target = getCurrentPendingTarget();
    elements.battleActionMessage.textContent = target
      ? `${target.displayName}: did the spell save pass or fail?`
      : "Resolve the spell save.";
    elements.battleTargetList.innerHTML = "";
    return;
  }

  if (mode === "damage") {
    const target = getCurrentPendingTarget();
    elements.battleActionMessage.textContent = target
      ? `Enter damage for ${target.displayName}. HP will not go below 0.`
      : "Enter damage.";
    elements.battleTargetList.innerHTML = "";
  }
}

function resetBattleActionUi(message) {
  elements.battleActionMessage.textContent = message;
  elements.battleTargetList.innerHTML = "";
  elements.battleResultButtons.classList.add("hidden");
  elements.battleSpellSaveButtons.classList.add("hidden");
  elements.damageForm.classList.add("hidden");
  elements.battleActionButtons.classList.remove("hidden");
}

function getCurrentPendingTarget() {
  const targetId = state.combat.pendingDamageTargets[0] || state.combat.selectedTargetId;
  return state.combat.combatants.find((entry) => entry.id === targetId) || null;
}

function getSelectedBattleTargets() {
  const targetIds = state.combat.selectedTargetIds.length
    ? state.combat.selectedTargetIds
    : state.combat.selectedTargetId
      ? [state.combat.selectedTargetId]
      : [];
  return targetIds
    .map((targetId) => state.combat.combatants.find((entry) => entry.id === targetId))
    .filter(Boolean);
}

function logBattleAction(text) {
  const active = getActiveCombatant();
  state.combat.actions.push({
    id: createId("action"),
    actorId: active?.id || "",
    actorName: active?.displayName || "Unknown",
    text,
    createdAt: new Date().toISOString(),
  });
}

function applyDamageToTarget(targetId, damage) {
  const target = state.combat.combatants.find((entry) => entry.id === targetId);
  if (!target) return;
  const amount = Math.max(0, Number(damage || 0));
  target.currentHealth = Math.max(0, Number(target.currentHealth || 0) - amount);
  target.defeated = target.currentHealth <= 0;
  logBattleAction(`${target.displayName} took ${amount} damage and has ${target.currentHealth}/${target.maxHealth} HP.`);
}

function getCombatantName(id) {
  return state.combat.combatants.find((entry) => entry.id === id)?.displayName || "Unknown target";
}

function completeBattle() {
  if (!state.currentSessionId || !state.combat.combatants.length) return;
  const session = getCurrentSession();
  if (!session) return;
  const endedAt = new Date().toISOString();
  const finalCombatants = state.combat.combatants.map((combatant) => ({ ...combatant }));
  const defeated = finalCombatants.filter((combatant) => combatant.defeated || combatant.currentHealth <= 0);
  const summary = `${finalCombatants.length} combatants. ${defeated.length} defeated.`;
  const battleRecord = {
    id: createId("battle"),
    name: `Battle - ${formatDate(session.date)}`,
    combatants: finalCombatants.map((combatant) => ({
      id: combatant.id,
      type: combatant.type,
      displayName: combatant.displayName,
      armorClass: combatant.armorClass,
      maxHealth: combatant.maxHealth,
      startingHealth: combatant.startingHealth ?? combatant.maxHealth,
      initiativeRoll: combatant.initiativeRoll,
    })),
    finalCombatants,
    activeIndex: state.combat.activeIndex,
    status: "completed",
    startedAt: state.combat.startedAt || endedAt,
    endedAt,
    actions: [...state.combat.actions],
    summary,
  };
  session.combats = [...(session.combats || []).filter((battle) => battle.id !== "current-combat"), battleRecord];
  finalCombatants
    .filter((combatant) => combatant.type === "character" && combatant.memberId)
    .forEach((combatant) => {
      const member = state.members.find((entry) => entry.id === combatant.memberId);
      if (member) {
        member.currentHealth = clampHealth(combatant.currentHealth, member.maxHealth);
      }
    });
  state.combat = clone(initialState.combat);
}

function sortInitiative() {
  state.combat.combatants.sort((a, b) => {
    const bRoll = b.initiativeRoll === "" ? -Infinity : Number(b.initiativeRoll);
    const aRoll = a.initiativeRoll === "" ? -Infinity : Number(a.initiativeRoll);
    return bRoll - aRoll;
  });
  state.combat.sorted = true;
  state.combat.activeIndex = 0;
  render();
}

function autoRollInitiative() {
  syncCombatants();
  state.combat.combatants = state.combat.combatants.map((combatant) => ({
    ...combatant,
    initiativeRoll: String(Math.floor(Math.random() * 20) + 1),
  }));
  sortInitiative();
}

function moveCombatant(id, direction) {
  const index = state.combat.combatants.findIndex((combatant) => combatant.id === id);
  const targetIndex = index + direction;
  if (index < 0 || targetIndex < 0 || targetIndex >= state.combat.combatants.length) return;
  const [combatant] = state.combat.combatants.splice(index, 1);
  state.combat.combatants.splice(targetIndex, 0, combatant);
  if (state.combat.activeIndex === index) {
    state.combat.activeIndex = targetIndex;
  } else if (state.combat.activeIndex === targetIndex) {
    state.combat.activeIndex = index;
  }
  state.combat.sorted = true;
  render();
}

function saveCombatToSession() {
  if (!state.currentSessionId || !state.combat.combatants.length) return;
  const session = state.sessions.find((entry) => entry.id === state.currentSessionId);
  if (!session) return;
  const currentCombat = {
    id: "current-combat",
    name: "Current battle",
    combatants: state.combat.combatants,
    activeIndex: state.combat.activeIndex,
    status: state.combat.sorted ? "active" : "setup",
    startedAt: state.combat.startedAt,
    actions: state.combat.actions,
    finalCombatants: state.combat.combatants,
    summary: "Battle in progress.",
  };
  session.participantMemberIds = state.combat.selectedMemberIds;
  session.combats = [...(session.combats || []).filter((battle) => battle.id !== "current-combat"), currentCombat];
}

function resetCombat() {
  state.combat = clone(initialState.combat);
  closeInitiativeScreen();
  render();
}

function closeInitiativeScreen() {
  elements.initiativeScreen.classList.add("hidden");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatDate(dateString) {
  const date = new Date(`${dateString}T00:00:00`);
  return new Intl.DateTimeFormat(undefined, { dateStyle: "medium" }).format(date);
}

function getActiveCampaign() {
  return state.campaigns.find((campaign) => campaign.id === state.currentCampaignId) || null;
}

function getCurrentSession() {
  return state.sessions.find((session) => session.id === state.currentSessionId) || null;
}

function getCampaignSessions(targetState, campaignId) {
  if (!campaignId) return [];
  return targetState.sessions
    .filter((session) => session.campaignId === campaignId)
    .sort((a, b) => new Date(a.date) - new Date(b.date) || (a.createdAt || 0) - (b.createdAt || 0));
}

function getCompletedBattleResults() {
  return getCampaignSessions(state, state.currentCampaignId).flatMap((session, sessionIndex) =>
    (session.combats || [])
      .filter((battle) => battle.id !== "current-combat" && battle.status === "completed")
      .map((battle) => ({ battle, session, sessionIndex })),
  );
}

function getCampaignDms(campaign) {
  const ids = campaign?.dmIds?.length ? campaign.dmIds : state.dms.map((dm) => dm.id);
  return ids.map((id) => state.dms.find((dm) => dm.id === id)).filter(Boolean);
}

function buildBattleSummaryMarkdown(battle) {
  const finalCombatants = battle.finalCombatants || [];
  const defeated = finalCombatants
    .filter((combatant) => combatant.defeated || Number(combatant.currentHealth) <= 0)
    .map((combatant) => combatant.displayName);
  const lines = [
    `### ${battle.name || "Battle"}`,
    `- Result: ${battle.summary || `${finalCombatants.length} combatants`}`,
    `- Defeated: ${defeated.length ? defeated.join(", ") : "None"}`,
    "- Final HP:",
    ...finalCombatants.map((combatant) => {
      const starting = combatant.startingHealth ?? combatant.maxHealth;
      return `  - ${combatant.displayName}: ${starting}/${combatant.maxHealth} to ${combatant.currentHealth}/${combatant.maxHealth}`;
    }),
  ];
  const actions = (battle.actions || []).slice(-6);
  if (actions.length) {
    lines.push("- Action notes:", ...actions.map((action) => `  - ${action.text}`));
  }
  return lines.join("\n");
}

function buildSessionMarkdown(session) {
  const campaign = state.campaigns.find((entry) => entry.id === session.campaignId);
  const recap = normalizeRecap(session.recap);
  const label = getSessionLabel(session, getSessionIndex(session));
  const speakers = session.speakerStats || [];
  const battles = (session.combats || []).filter((battle) => battle.id !== "current-combat" && battle.status === "completed");
  return [
    `# ${label}`,
    "",
    `Campaign: ${campaign?.name || "Unknown campaign"}`,
    `Date: ${formatDate(session.date)}`,
    "",
    "## Recap",
    recap.summary || "No recap saved.",
    "",
    "## Important Events",
    recap.events || "None saved.",
    "",
    "## NPCs",
    recap.npcs || "None saved.",
    "",
    "## Locations",
    recap.locations || "None saved.",
    "",
    "## Quests",
    recap.quests || "None saved.",
    "",
    "## Unresolved Threads",
    recap.unresolvedThreads || "None saved.",
    "",
    "## Speaking Time",
    speakers.length ? speakers.map((speaker) => `- ${speaker.name}: ${speaker.minutes} min`).join("\n") : "No speaking time saved.",
    "",
    "## Battle Results",
    battles.length ? battles.map(buildBattleSummaryMarkdown).join("\n\n") : "No battle results saved.",
    "",
    "## Transcript",
    session.transcript || "No transcript saved.",
    "",
  ].join("\n");
}

function buildCampaignMarkdown(campaign) {
  const sessions = getCampaignSessions(state, campaign.id);
  return [
    `# ${campaign.name}`,
    "",
    `Dungeon Masters: ${formatNameList(getCampaignDms(campaign).map((dm) => dm.name)) || "None saved"}`,
    `Sessions: ${sessions.length}`,
    "",
    ...sessions.map(buildSessionMarkdown),
  ].join("\n\n");
}

function downloadTextFile(fileName, content) {
  const blob = new Blob([content], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function slugifyFileName(value) {
  return String(value || "dnd-notes")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80) || "dnd-notes";
}

function getCampaignDisplayTitle(campaign) {
  const dms = getCampaignDms(campaign).map((dm) => dm.name);
  if (dms.length) {
    return `${formatNameList(dms)}'s Campaign`;
  }
  return campaign?.name || "D&D Club Campaign";
}

function formatNameList(names) {
  if (names.length <= 1) return names[0] || "";
  if (names.length === 2) return `${names[0]} and ${names[1]}`;
  return `${names.slice(0, -1).join(", ")}, and ${names[names.length - 1]}`;
}

function getSessionLabel(session, index) {
  return `Session ${index + 1} - ${formatDate(session.date)}`;
}

function getSessionIndex(session) {
  return Math.max(0, getCampaignSessions(state, session?.campaignId).findIndex((entry) => entry.id === session?.id));
}

function getSpeakers() {
  const players = state.members.map((member) => ({
    id: member.id,
    type: "member",
    name: `${member.playerName} (${member.characterName})`,
  }));
  const dms = state.dms.map((dm) => ({
    id: dm.id,
    type: "dm",
    name: dm.name,
  }));
  return [...players, ...dms];
}

function formatMinutes(minutes) {
  const value = Math.max(0, Number(minutes || 0));
  return Number.isInteger(value) ? String(value) : String(Math.round(value * 10) / 10);
}

function getUploadedAudioFile() {
  const file = elements.audioUploadFile.files?.[0];
  if (!file) {
    elements.audioUploadMessage.textContent = "Choose a downloaded recording file first.";
    return null;
  }

  const fileName = file.name || "session-audio";
  const extension = `.${fileName.split(".").pop().toLowerCase()}`;
  if (!SUPPORTED_AUDIO_EXTENSIONS.includes(extension)) {
    elements.audioUploadMessage.textContent = "Use mp3, mp4, mpeg, mpga, m4a, wav, or webm audio.";
    audioUploadStatus = "error";
    render();
    return null;
  }

  if (file.size > MAX_AUDIO_UPLOAD_BYTES) {
    elements.audioUploadMessage.textContent = "Audio files must be 25 MB or smaller for this audio processing step.";
    audioUploadStatus = "error";
    render();
    return null;
  }

  return file;
}

function normalizeSpeakerTimingSegment(segment, index) {
  const startSeconds = Math.max(0, Number(segment.startSeconds ?? segment.start ?? 0) || 0);
  const endSeconds = Math.max(startSeconds, Number(segment.endSeconds ?? segment.end ?? startSeconds) || startSeconds);
  const durationSeconds = Math.max(
    0,
    Number(segment.durationSeconds ?? segment.duration ?? endSeconds - startSeconds) || 0,
  );
  return {
    id: segment.id || createId("segment"),
    speakerLabel: String(segment.speakerLabel || segment.speaker || `Speaker ${index + 1}`).trim(),
    assignedSpeakerId: segment.assignedSpeakerId || "",
    assignedSpeakerType: segment.assignedSpeakerType || "",
    assignedSpeakerName: segment.assignedSpeakerName || "",
    minutes: Math.round((durationSeconds / 60) * 10) / 10,
    durationSeconds,
    startSeconds,
    endSeconds,
    note: String(segment.text || segment.note || "").trim(),
    source: "ai",
    createdAt: segment.createdAt || new Date().toISOString(),
  };
}

function formatDuration(totalSeconds) {
  const seconds = Math.max(0, Math.round(totalSeconds || 0));
  const minutes = Math.floor(seconds / 60);
  const remainder = seconds % 60;
  return `${minutes}:${String(remainder).padStart(2, "0")}`;
}

function formatStatus(status) {
  return String(status || "not-ready")
    .split("-")
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

function getAudioProcessingStatus(session) {
  if (!session) return "Not uploaded";
  if (state.settings.manualMode) return "Manual Mode";
  if (speakerTimingStatus === "processing") return "Processing speakers";
  if (audioUploadStatus === "uploading") return "Uploading";
  if (audioUploadStatus === "transcribing") return "Transcribing";
  if (audioUploadStatus === "error" || speakerTimingStatus === "error") return "Error";
  if (session.speakerSegments?.some((segment) => segment.source === "ai")) return "Needs review";
  if (session.recording?.status === "transcribed") return "Complete";
  if (session.recording?.status === "speaker timing ready") return "Needs review";
  if (session.recording?.status === "saved") return "Ready for backend";
  if (session.recording?.status === "ready to download") return "Ready";
  if (session.recording?.status === "recording") return "Recording";
  return "Not uploaded";
}

function getAudioProcessingMessage(session) {
  if (!session) return "Create or select a session before preparing audio.";
  if (state.settings.manualMode) {
    return "Manual Mode is on. Record and download audio, then paste transcripts manually until an AI backend is connected.";
  }
  if (!elements.audioUploadFile.files?.[0]) {
    return "Choose a downloaded audio file before transcription or AI speaker timing.";
  }
  return `Upload a saved recording to ${API_BASE_URL} for transcription or AI speaker timing. Speaker labels are temporary until you review them.`;
}

function updateManualModeButton(button) {
  const enabled = Boolean(state.settings.manualMode);
  button.setAttribute("aria-pressed", String(enabled));
  button.textContent = enabled ? "Manual Mode: On" : "Manual Mode: Off";
}

function getAiToolsMessage(session) {
  if (state.settings.manualMode) {
    return "Manual Mode is on. Turn it off to use the deployed AI backend, or keep writing notes manually.";
  }
  if (session.aiStatus === "processing") return "Generating a structured recap from the transcript.";
  if (session.aiStatus === "complete") return "AI recap fields were generated. Review them, then save the notes.";
  if (session.aiStatus === "error") return "AI recap generation failed. Existing notes were kept, and you can try again.";
  return `AI recap generation sends the transcript to ${API_BASE_URL}. The public website uses the deployed Render backend by default.`;
}

function updateGenerateRecapButton() {
  if (!elements.generateRecap) return;
  const session = getCurrentSession();
  const transcript = elements.transcriptField?.value.trim() || "";
  const disabled = state.settings.manualMode || !session || !transcript || session.aiStatus === "processing";
  elements.generateRecap.disabled = disabled;
  if (state.settings.manualMode) {
    elements.generateRecap.textContent = "Manual Mode enabled";
  } else if (!session) {
    elements.generateRecap.textContent = "Generate recap from transcript";
  } else if (session.aiStatus === "processing") {
    elements.generateRecap.textContent = "Generating recap...";
  } else if (!transcript) {
    elements.generateRecap.textContent = "Add transcript to generate recap";
  } else {
    elements.generateRecap.textContent = "Generate recap from transcript";
  }
}

function renderAiSetupStatus(session) {
  const setup = state.settings.aiSetup || clone(initialState.settings.aiSetup);
  const hasBackendUrl = Boolean(API_BASE_URL);
  const backendReady = setup.backendReachable === true;
  const keyReady = setup.hasOpenAIKey === true;
  const manualMode = Boolean(state.settings.manualMode);
  const transcriptReady = Boolean(session && (elements.transcriptField?.value || session.transcript || "").trim());
  const aiReady = hasBackendUrl && backendReady && keyReady && !manualMode;
  const recapReady = aiReady && transcriptReady;
  const transcriptionReady = aiReady && Boolean(session);

  elements.aiSetupStatus.textContent = backendReady && keyReady ? "Connected" : backendReady ? "Needs key" : "Not checked";
  elements.aiSetupMessage.textContent = getAiSetupMessage(setup, {
    manualMode,
    hasBackendUrl,
    backendReady,
    keyReady,
    transcriptReady,
    sessionReady: Boolean(session),
  });
  elements.aiSetupManualMode.textContent = manualMode ? "On - AI disabled" : "Off - AI allowed";
  elements.aiSetupBackendUrl.textContent = hasBackendUrl ? "Saved" : "Missing";
  elements.aiSetupBackendHealth.textContent =
    setup.backendReachable === true ? "Connected" : setup.backendReachable === false ? "Unreachable" : "Not checked";
  elements.aiSetupOpenAiKey.textContent =
    setup.hasOpenAIKey === true ? "Connected" : setup.hasOpenAIKey === false ? "Missing" : "Unknown";
  elements.aiSetupRecapReady.textContent = recapReady
    ? "Ready"
    : !session
      ? "Needs session"
      : !transcriptReady
        ? "Needs transcript"
        : !aiReady
          ? "Not ready"
          : "Ready";
  elements.aiSetupTranscriptionReady.textContent = transcriptionReady
    ? "Ready after choosing audio"
    : !session
      ? "Needs session"
      : "Not ready";
  elements.aiSetupCheckedAt.textContent = setup.lastCheckedAt
    ? `Last checked: ${new Date(setup.lastCheckedAt).toLocaleString()}`
    : "Last checked: never";
  elements.checkAiSetup.disabled = setup.message === "Checking AI setup...";
}

function getAiSetupMessage(setup, details) {
  if (setup.message === "Checking AI setup...") return setup.message;
  if (!details.hasBackendUrl) return "Add a backend URL before checking AI setup.";
  if (setup.backendReachable === false) return setup.message || "Backend could not be reached.";
  if (setup.hasOpenAIKey === false) return "Backend is reachable, but the OpenAI API key is missing on Render.";
  if (details.manualMode && details.backendReady && details.keyReady) {
    return "Backend and OpenAI are connected. Turn Manual Mode off to use AI features.";
  }
  if (details.backendReady && details.keyReady) return "Backend and OpenAI are connected.";
  return setup.message || "Check the backend before using AI recap or transcription.";
}

function getCurrentRecapDraft() {
  return {
    summary: elements.recapSummaryField.value.trim(),
    events: elements.recapEventsField.value.trim(),
    npcs: elements.recapNpcsField.value.trim(),
    locations: elements.recapLocationsField.value.trim(),
    quests: elements.recapQuestsField.value.trim(),
    unresolvedThreads: elements.recapThreadsField.value.trim(),
  };
}

function updateRecordingDuration() {
  if (!recordingStartedAt) return;
  elements.recordingDuration.textContent = formatDuration((Date.now() - recordingStartedAt) / 1000);
}

function getAudioMimeType() {
  if (!window.MediaRecorder) return "";
  if (MediaRecorder.isTypeSupported("audio/webm")) return "audio/webm";
  if (MediaRecorder.isTypeSupported("audio/mp4")) return "audio/mp4";
  return "";
}

function getAudioExtension(mimeType) {
  return mimeType.includes("mp4") ? "m4a" : "webm";
}

elements.navButtons.forEach((button) => {
  button.addEventListener("click", () => showView(button.dataset.view));
});

elements.memberForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const maxHealth = Number(elements.maxHealth.value);
  const currentHealth = clampHealth(elements.currentHealth.value, maxHealth);
  const member = {
    id: elements.memberId.value || createId("member"),
    playerName: elements.playerName.value.trim(),
    characterName: elements.characterName.value.trim(),
    characterRace: elements.characterRace.value.trim(),
    armorClass: Number(elements.armorClass.value),
    maxHealth,
    currentHealth,
  };
  if (
    !member.playerName ||
    !member.characterName ||
    !member.characterRace ||
    !member.armorClass ||
    !member.maxHealth ||
    elements.currentHealth.value === ""
  )
    return;

  const existingIndex = state.members.findIndex((entry) => entry.id === member.id);
  if (existingIndex >= 0) {
    state.members[existingIndex] = { ...state.members[existingIndex], ...member };
  } else {
    state.members.push({ ...member, voiceSampleId: null });
  }

  elements.memberForm.reset();
  elements.memberId.value = "";
  elements.memberFormTitle.textContent = "Add a player character";
  elements.cancelEdit.classList.add("hidden");
  render();
});

elements.cancelEdit.addEventListener("click", () => {
  elements.memberForm.reset();
  elements.memberId.value = "";
  elements.memberFormTitle.textContent = "Add a player character";
  elements.cancelEdit.classList.add("hidden");
});

elements.memberList.addEventListener("click", (event) => {
  const editId = event.target.dataset.editMember;
  const deleteId = event.target.dataset.deleteMember;

  if (editId) {
    const member = state.members.find((entry) => entry.id === editId);
    if (!member) return;
    elements.memberId.value = member.id;
    elements.playerName.value = member.playerName;
    elements.characterName.value = member.characterName;
    elements.characterRace.value = member.characterRace;
    elements.armorClass.value = member.armorClass || "";
    elements.maxHealth.value = member.maxHealth || "";
    elements.currentHealth.value = member.currentHealth ?? member.maxHealth ?? "";
    elements.memberFormTitle.textContent = "Edit player character";
    elements.cancelEdit.classList.remove("hidden");
  }

  if (deleteId) {
    state.members = state.members.filter((member) => member.id !== deleteId);
    state.combat.selectedMemberIds = state.combat.selectedMemberIds.filter((id) => id !== deleteId);
    state.combat.combatants = state.combat.combatants.filter((combatant) => combatant.memberId !== deleteId);
    render();
  }
});

elements.dmForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const dm = {
    id: elements.dmId.value || createId("dm"),
    name: elements.dmName.value.trim(),
  };
  if (!dm.name) return;

  const existingIndex = state.dms.findIndex((entry) => entry.id === dm.id);
  if (existingIndex >= 0) {
    state.dms[existingIndex] = { ...state.dms[existingIndex], ...dm };
  } else {
    state.dms.push({ ...dm, voiceSampleId: null });
    state.campaigns = state.campaigns.map((campaign) =>
      campaign.status === "active" ? { ...campaign, dmIds: [...new Set([...(campaign.dmIds || []), dm.id])] } : campaign,
    );
  }

  elements.dmForm.reset();
  elements.dmId.value = "";
  elements.dmFormTitle.textContent = "Add a Dungeon Master";
  elements.cancelDmEdit.classList.add("hidden");
  render();
});

elements.cancelDmEdit.addEventListener("click", () => {
  elements.dmForm.reset();
  elements.dmId.value = "";
  elements.dmFormTitle.textContent = "Add a Dungeon Master";
  elements.cancelDmEdit.classList.add("hidden");
});

elements.dmList.addEventListener("click", (event) => {
  const editId = event.target.dataset.editDm;
  const deleteId = event.target.dataset.deleteDm;

  if (editId) {
    const dm = state.dms.find((entry) => entry.id === editId);
    if (!dm) return;
    elements.dmId.value = dm.id;
    elements.dmName.value = dm.name;
    elements.dmFormTitle.textContent = "Edit Dungeon Master";
    elements.cancelDmEdit.classList.remove("hidden");
  }

  if (deleteId) {
    state.dms = state.dms.filter((dm) => dm.id !== deleteId);
    state.campaigns = state.campaigns.map((campaign) => ({
      ...campaign,
      dmIds: (campaign.dmIds || []).filter((dmId) => dmId !== deleteId),
    }));
    render();
  }
});

elements.combatMemberOptions.addEventListener("change", (event) => {
  const memberId = event.target.dataset.combatMember;
  if (!memberId) return;
  if (event.target.checked) {
    state.combat.selectedMemberIds.push(memberId);
  } else {
    state.combat.selectedMemberIds = state.combat.selectedMemberIds.filter((id) => id !== memberId);
  }
  state.combat.sorted = false;
  closeInitiativeScreen();
  render();
});

elements.monsterForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const type = elements.monsterType.value.trim();
  const count = Math.max(1, Number(elements.monsterCount.value || 1));
  const armorClass = Number(elements.monsterArmorClass.value || 0);
  const maxHealth = Number(elements.monsterHealth.value || 0);
  if (!type || !armorClass || !maxHealth) return;
  for (let index = 1; index <= count; index += 1) {
    state.combat.monsters.push({
      id: createId("monster"),
      type,
      displayName: count === 1 ? type : `${type} ${index}`,
      armorClass,
      maxHealth,
      currentHealth: maxHealth,
      defeated: false,
    });
  }
  elements.monsterType.value = "";
  elements.monsterCount.value = 1;
  elements.monsterArmorClass.value = 10;
  elements.monsterHealth.value = 10;
  state.combat.sorted = false;
  closeInitiativeScreen();
  render();
});

elements.monsterList.addEventListener("click", (event) => {
  const monsterId = event.target.dataset.deleteMonster;
  if (!monsterId) return;
  state.combat.monsters = state.combat.monsters.filter((monster) => monster.id !== monsterId);
  state.combat.combatants = state.combat.combatants.filter((combatant) => combatant.id !== monsterId);
  closeInitiativeScreen();
  render();
});

elements.combatantList.addEventListener("input", (event) => {
  const combatantId = event.target.dataset.roll;
  if (!combatantId) return;
  const combatant = state.combat.combatants.find((entry) => entry.id === combatantId);
  if (combatant) {
    combatant.initiativeRoll = event.target.value;
    state.combat.sorted = false;
    closeInitiativeScreen();
    renderCurrentTurn();
    renderInitiativeScreen();
    saveState();
  }
});

elements.combatantList.addEventListener("click", (event) => {
  const upId = event.target.dataset.moveUp;
  const downId = event.target.dataset.moveDown;
  if (upId) moveCombatant(upId, -1);
  if (downId) moveCombatant(downId, 1);
});

elements.autoRollInitiative.addEventListener("click", autoRollInitiative);
elements.sortInitiative.addEventListener("click", sortInitiative);
elements.resetCombat.addEventListener("click", resetCombat);

elements.startInitiative.addEventListener("click", () => {
  if (!state.combat.combatants.length || !state.combat.sorted || !hasCompleteInitiative()) return;
  state.combat.startedAt = state.combat.startedAt || new Date().toISOString();
  state.combat.battleRunning = true;
  ensureActiveLivingCombatant();
  elements.initiativeScreen.classList.remove("hidden");
  renderInitiativeScreen();
  saveState();
});

elements.exitInitiativeScreen.addEventListener("click", () => {
  completeBattle();
  closeInitiativeScreen();
  render();
});

elements.battleOther.addEventListener("click", () => {
  logBattleAction("Turn ended with Other.");
  advanceTurn();
});

elements.battleAttack.addEventListener("click", () => {
  state.combat.actionMode = "attack-target";
  state.combat.selectedTargetId = null;
  state.combat.selectedTargetIds = [];
  renderInitiativeScreen();
});

elements.battleMultiAttack.addEventListener("click", () => {
  state.combat.actionMode = "multi-target";
  state.combat.selectedTargetId = null;
  state.combat.selectedTargetIds = [];
  renderInitiativeScreen();
});

elements.battleTargetList.addEventListener("change", (event) => {
  const targetId = event.target.dataset.battleTarget;
  if (!targetId) return;
  if (state.combat.actionMode === "multi-target") {
    if (event.target.checked) {
      state.combat.selectedTargetIds = [...new Set([...state.combat.selectedTargetIds, targetId])];
    } else {
      state.combat.selectedTargetIds = state.combat.selectedTargetIds.filter((id) => id !== targetId);
    }
  } else {
    state.combat.selectedTargetId = targetId;
    state.combat.actionMode = "target-selected";
  }
  renderInitiativeScreen();
});

elements.battleTargetList.addEventListener("click", (event) => {
  if (event.target.id !== "confirm-multi-targets") return;
  if (!state.combat.selectedTargetIds.length) return;
  state.combat.actionMode = "target-selected";
  logBattleAction(`Multi-attack targets: ${state.combat.selectedTargetIds.map(getCombatantName).join(", ")}.`);
  renderInitiativeScreen();
});

elements.battlePass.addEventListener("click", () => {
  const targets = getSelectedBattleTargets();
  if (!targets.length) return;
  state.combat.pendingDamageTargets = targets.map((target) => target.id);
  state.combat.damageMode = "attack";
  state.combat.actionMode = "damage";
  logBattleAction(`${targets.length > 1 ? "Multi-attack" : "Attack"} passed against ${targets.map((target) => target.displayName).join(", ")}.`);
  renderInitiativeScreen();
});

elements.battleFail.addEventListener("click", () => {
  const targets = getSelectedBattleTargets();
  logBattleAction(`${targets.length > 1 ? "Multi-attack" : "Attack"} failed against ${targets.map((target) => target.displayName).join(", ")}.`);
  advanceTurn();
});

elements.battleSpell.addEventListener("click", () => {
  const targets = getSelectedBattleTargets();
  if (!targets.length) return;
  state.combat.pendingDamageTargets = targets.map((target) => target.id);
  state.combat.damageMode = "spell";
  state.combat.actionMode = "spell-save";
  renderInitiativeScreen();
});

elements.spellSavePass.addEventListener("click", () => {
  const target = getCurrentPendingTarget();
  if (!target) return;
  logBattleAction(`Spell save passed for ${target.displayName}.`);
  state.combat.actionMode = "damage";
  renderInitiativeScreen();
});

elements.spellSaveFail.addEventListener("click", () => {
  const target = getCurrentPendingTarget();
  if (target) logBattleAction(`Spell save failed for ${target.displayName}.`);
  state.combat.pendingDamageTargets.shift();
  if (state.combat.pendingDamageTargets.length) {
    state.combat.actionMode = "spell-save";
    renderInitiativeScreen();
  } else {
    advanceTurn();
  }
});

elements.damageForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const target = getCurrentPendingTarget();
  if (!target) return;
  applyDamageToTarget(target.id, elements.damageAmount.value);
  elements.damageAmount.value = 0;
  state.combat.pendingDamageTargets.shift();
  if (state.combat.pendingDamageTargets.length) {
    state.combat.actionMode = state.combat.damageMode === "spell" ? "spell-save" : "damage";
    render();
  } else {
    advanceTurn();
  }
});

elements.campaignForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const campaignName = elements.campaignName.value.trim();
  if (!campaignName) return;

  if (elements.campaignId.value) {
    state.campaigns = state.campaigns.map((campaign) =>
      campaign.id === elements.campaignId.value ? { ...campaign, name: campaignName } : campaign,
    );
  } else {
    const campaign = {
      id: createId("campaign"),
      name: campaignName,
      dmIds: state.dms.map((dm) => dm.id),
      sessionIds: [],
      status: "active",
    };
    state.campaigns = state.campaigns.map((entry) => ({ ...entry, status: "finished" }));
    state.campaigns.unshift(campaign);
    state.currentCampaignId = campaign.id;
    state.currentSessionId = null;
  }
  elements.campaignId.value = "";
  elements.campaignName.value = "";
  elements.campaignFormTitle.textContent = "Create a campaign";
  elements.cancelCampaignEdit.classList.add("hidden");
  render();
});

elements.cancelCampaignEdit.addEventListener("click", () => {
  elements.campaignId.value = "";
  elements.campaignName.value = "";
  elements.campaignFormTitle.textContent = "Create a campaign";
  elements.cancelCampaignEdit.classList.add("hidden");
});

elements.sessionForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const campaign = getActiveCampaign();
  if (!campaign) return;
  const session = {
    id: createId("session"),
    date: elements.sessionDate.value,
    campaignId: campaign.id,
    participantMemberIds: [],
    combats: [],
    recording: null,
    transcript: null,
    speakerStats: null,
    aiStatus: "not-ready",
    speakerSegments: [],
    speakerReviewStatus: "manual",
    recap: null,
    createdAt: Date.now(),
  };
  if (!session.date) return;
  state.sessions.push(session);
  campaign.sessionIds = [...new Set([...(campaign.sessionIds || []), session.id])];
  state.currentSessionId = session.id;
  elements.sessionDate.valueAsDate = new Date();
  render();
});

elements.campaignList.addEventListener("click", (event) => {
  const currentId = event.target.dataset.currentCampaign;
  const finishId = event.target.dataset.finishCampaign;
  const editId = event.target.dataset.editCampaign;
  const deleteId = event.target.dataset.deleteCampaign;

  if (currentId) {
    state.campaigns = state.campaigns.map((campaign) =>
      campaign.id === currentId ? { ...campaign, status: "active" } : { ...campaign, status: "finished" },
    );
    state.currentCampaignId = currentId;
    state.currentSessionId = getCampaignSessions(state, currentId)[0]?.id || null;
    render();
  }

  if (finishId) {
    const campaignToToggle = state.campaigns.find((campaign) => campaign.id === finishId);
    const reopening = campaignToToggle?.status === "finished";
    state.campaigns = state.campaigns.map((campaign) => {
      if (campaign.id === finishId) {
        return { ...campaign, status: reopening ? "active" : "finished" };
      }
      return reopening ? { ...campaign, status: "finished" } : campaign;
    });
    if (reopening) {
      state.currentCampaignId = finishId;
      state.currentSessionId = getCampaignSessions(state, finishId)[0]?.id || null;
    } else if (state.currentCampaignId === finishId) {
      state.currentCampaignId = state.campaigns.find((campaign) => campaign.status === "active")?.id || finishId;
      state.currentSessionId = getCampaignSessions(state, state.currentCampaignId)[0]?.id || null;
    }
    render();
  }

  if (editId) {
    const campaign = state.campaigns.find((entry) => entry.id === editId);
    if (!campaign) return;
    elements.campaignId.value = campaign.id;
    elements.campaignName.value = campaign.name;
    elements.campaignFormTitle.textContent = "Edit campaign";
    elements.cancelCampaignEdit.classList.remove("hidden");
  }

  if (deleteId) {
    state.campaigns = state.campaigns.filter((campaign) => campaign.id !== deleteId);
    state.sessions = state.sessions.filter((session) => session.campaignId !== deleteId);
    if (elements.campaignId.value === deleteId) {
      elements.campaignId.value = "";
      elements.campaignName.value = "";
      elements.campaignFormTitle.textContent = "Create a campaign";
      elements.cancelCampaignEdit.classList.add("hidden");
    }
    if (state.currentCampaignId === deleteId) {
      state.currentCampaignId =
        state.campaigns.find((campaign) => campaign.status === "active")?.id || state.campaigns[0]?.id || null;
      state.currentSessionId = getCampaignSessions(state, state.currentCampaignId)[0]?.id || null;
    }
    render();
  }
});

elements.sessionList.addEventListener("click", (event) => {
  const currentId = event.target.dataset.currentSession;
  const deleteId = event.target.dataset.deleteSession;
  const viewBattleId = event.target.dataset.viewBattle;
  const battleSessionId = event.target.dataset.battleSession;
  if (viewBattleId && battleSessionId) {
    state.currentSessionId = battleSessionId;
    highlightedBattleId = viewBattleId;
    showView("initiative");
    render();
    document.querySelector(`#battle-result-${CSS.escape(viewBattleId)}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }
  if (currentId) {
    state.currentSessionId = currentId;
    render();
  }
  if (deleteId) {
    state.sessions = state.sessions.filter((session) => session.id !== deleteId);
    state.campaigns = state.campaigns.map((campaign) => ({
      ...campaign,
      sessionIds: (campaign.sessionIds || []).filter((sessionId) => sessionId !== deleteId),
    }));
    if (state.currentSessionId === deleteId) {
      state.currentSessionId = getCampaignSessions(state, state.currentCampaignId)[0]?.id || null;
    }
    render();
  }
});

elements.battleResultsList.addEventListener("click", (event) => {
  const copyBattleId = event.target.dataset.copyBattle;
  const battleId = event.target.dataset.deleteBattle;
  const sessionId = event.target.dataset.battleSession;
  if (copyBattleId && sessionId) {
    const session = state.sessions.find((entry) => entry.id === sessionId);
    const battle = session?.combats?.find((entry) => entry.id === copyBattleId);
    if (!session || !battle) return;
    const recap = normalizeRecap(session.recap);
    const summary = buildBattleSummaryMarkdown(battle);
    session.recap = {
      ...recap,
      events: [recap.events, summary].filter(Boolean).join("\n\n"),
    };
    render();
    showView("progress");
    return;
  }
  if (!battleId || !sessionId) return;
  const session = state.sessions.find((entry) => entry.id === sessionId);
  if (!session) return;
  session.combats = (session.combats || []).filter((battle) => battle.id !== battleId);
  if (highlightedBattleId === battleId) highlightedBattleId = null;
  render();
});

elements.startRecording.addEventListener("click", async () => {
  const session = getCurrentSession();
  if (!session) return;
  if (!navigator.mediaDevices?.getUserMedia || !window.MediaRecorder) {
    elements.recordingMessage.textContent = "This browser does not support local microphone recording.";
    return;
  }

  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    const mimeType = getAudioMimeType();
    recordingChunks = [];
    recordingStartedAt = Date.now();
    mediaRecorder = new MediaRecorder(stream, mimeType ? { mimeType } : undefined);
    session.recording = {
      status: "recording",
      startedAt: new Date(recordingStartedAt).toISOString(),
      durationSeconds: 0,
      fileName: "",
    };
    elements.recordingMessage.textContent = "Recording in progress.";
    if (recordingDownloadUrl) {
      URL.revokeObjectURL(recordingDownloadUrl);
      recordingDownloadUrl = null;
    }
    elements.downloadRecording.classList.add("hidden");

    mediaRecorder.addEventListener("dataavailable", (event) => {
      if (event.data.size > 0) recordingChunks.push(event.data);
    });

    mediaRecorder.addEventListener("stop", () => {
      const durationSeconds = Math.round((Date.now() - recordingStartedAt) / 1000);
      const type = mediaRecorder.mimeType || mimeType || "audio/webm";
      const extension = getAudioExtension(type);
      const fileName = `dnd-session-${session.date}.${extension}`;
      const blob = new Blob(recordingChunks, { type });
      recordingDownloadUrl = URL.createObjectURL(blob);
      session.recording = {
        status: "ready to download",
        startedAt: session.recording?.startedAt || new Date(recordingStartedAt).toISOString(),
        durationSeconds,
        fileName,
      };
      elements.downloadRecording.href = recordingDownloadUrl;
      elements.downloadRecording.download = fileName;
      elements.downloadRecording.classList.remove("hidden");
      elements.recordingMessage.textContent = "Recording stopped. Download the audio file before closing this page.";
      stream.getTracks().forEach((track) => track.stop());
      clearInterval(recordingTimerId);
      mediaRecorder = null;
      recordingStartedAt = null;
      render();
    });

    mediaRecorder.start();
    recordingTimerId = setInterval(updateRecordingDuration, 500);
    render();
  } catch (error) {
    elements.recordingMessage.textContent = "Microphone access was blocked or unavailable. Check browser permissions and try again.";
    mediaRecorder = null;
    recordingStartedAt = null;
    clearInterval(recordingTimerId);
  }
});

elements.stopRecording.addEventListener("click", () => {
  if (mediaRecorder && mediaRecorder.state !== "inactive") {
    mediaRecorder.stop();
  }
});

function setManualMode(enabled) {
  state.settings.manualMode = Boolean(enabled);
  if (state.settings.manualMode) {
    audioUploadStatus = "ready";
    speakerTimingStatus = "ready";
  }
  render();
}

elements.speakingManualMode.addEventListener("click", () => {
  setManualMode(!state.settings.manualMode);
});

elements.progressManualMode.addEventListener("click", () => {
  setManualMode(!state.settings.manualMode);
});

function openSessionNotes() {
  showView("progress");
  renderProgressNotes();
}

elements.recordingPasteTranscript.addEventListener("click", () => {
  elements.manualTranscriptField.focus();
});

elements.recordingOpenNotes.addEventListener("click", openSessionNotes);
elements.manualTranscriptOpenNotes.addEventListener("click", openSessionNotes);

elements.saveManualTranscript.addEventListener("click", () => {
  const session = getCurrentSession();
  if (!session) return;
  session.transcript = elements.manualTranscriptField.value.trim();
  session.aiStatus = session.transcript ? "ready" : "not-ready";
  elements.manualTranscriptMessage.textContent = session.transcript
    ? "Transcript saved to Session Notes."
    : "Transcript cleared for this session.";
  render();
});

elements.audioUploadFile.addEventListener("change", () => {
  renderSpeakingTime();
});

elements.transcribeAudio.addEventListener("click", async () => {
  const session = getCurrentSession();
  if (!session) return;
  speakerTimingStatus = "ready";
  if (state.settings.manualMode) {
    elements.audioUploadMessage.textContent = "Manual Mode is on. Paste a transcript manually or turn Manual Mode off to use AI processing.";
    return;
  }

  const file = elements.audioUploadFile.files?.[0];
  if (!file) {
    elements.audioUploadMessage.textContent = "Choose a downloaded recording file first.";
    return;
  }

  const fileName = file.name || "session-audio";
  const extension = `.${fileName.split(".").pop().toLowerCase()}`;
  if (!SUPPORTED_AUDIO_EXTENSIONS.includes(extension)) {
    elements.audioUploadMessage.textContent = "Use mp3, mp4, mpeg, mpga, m4a, wav, or webm audio.";
    audioUploadStatus = "error";
    render();
    return;
  }

  if (file.size > MAX_AUDIO_UPLOAD_BYTES) {
    elements.audioUploadMessage.textContent = "Audio files must be 25 MB or smaller for this transcription step.";
    audioUploadStatus = "error";
    render();
    return;
  }

  const activeCampaign = getActiveCampaign();
  const sessionLabel = getSessionLabel(session, getSessionIndex(session));
  const formData = new FormData();
  formData.append("audio", file);
  formData.append("campaignName", activeCampaign?.name || "Active campaign");
  formData.append("sessionLabel", sessionLabel);

  speakerTimingStatus = "ready";
  audioUploadStatus = "uploading";
  elements.audioUploadMessage.textContent = "Uploading audio for transcription.";
  render();

  try {
    audioUploadStatus = "transcribing";
    elements.audioUploadMessage.textContent = "Transcribing audio. This can take a little while.";
    renderSpeakingTime();

    const response = await fetch(`${API_BASE_URL}/api/transcribe-session-audio`, {
      method: "POST",
      body: formData,
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || "Audio transcription failed.");
    }

    const transcript = String(payload.transcript || "").trim();
    if (!transcript) {
      throw new Error("The transcription finished, but no transcript text came back.");
    }

    session.transcript = transcript;
    session.aiStatus = "ready";
    session.recording = {
      ...(normalizeRecording(session.recording) || {}),
      status: "transcribed",
      durationSeconds: Number(payload.durationSeconds || session.recording?.durationSeconds || 0),
      fileName: payload.fileName || fileName,
    };
    audioUploadStatus = "complete";
    speakerTimingStatus = "ready";
    elements.audioUploadFile.value = "";
    elements.audioUploadMessage.textContent = "Transcript added to Session Notes. Review it, then save notes.";
    render();
  } catch (error) {
    audioUploadStatus = "error";
    elements.audioUploadMessage.textContent =
      error.message === "Failed to fetch"
        ? `Could not reach the AI backend at ${API_BASE_URL}.`
        : error.message || "Audio transcription failed.";
    renderSpeakingTime();
    saveState();
  }
});

elements.processSpeakerTiming.addEventListener("click", async () => {
  const session = getCurrentSession();
  if (!session) return;
  audioUploadStatus = "ready";
  if (state.settings.manualMode) {
    elements.audioUploadMessage.textContent = "Manual Mode is on. Turn it off to use AI speaker timing.";
    return;
  }

  const file = getUploadedAudioFile();
  if (!file) return;

  const activeCampaign = getActiveCampaign();
  const sessionLabel = getSessionLabel(session, getSessionIndex(session));
  const formData = new FormData();
  formData.append("audio", file);
  formData.append("campaignName", activeCampaign?.name || "Active campaign");
  formData.append("sessionLabel", sessionLabel);

  audioUploadStatus = "ready";
  speakerTimingStatus = "processing";
  elements.audioUploadMessage.textContent = "Processing speaker timing. This can take a little while.";
  renderSpeakingTime();

  try {
    const response = await fetch(`${API_BASE_URL}/api/process-speaker-timing`, {
      method: "POST",
      body: formData,
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || "Speaker timing failed.");
    }

    const transcript = String(payload.transcript || "").trim();
    const segments = Array.isArray(payload.speakerSegments)
      ? payload.speakerSegments.map(normalizeSpeakerTimingSegment)
      : [];

    if (transcript) {
      session.transcript = transcript;
      session.aiStatus = "ready";
    }
    session.speakerSegments = segments;
    session.speakerReviewStatus = segments.length ? "needs-review" : "manual";
    session.recording = {
      ...(normalizeRecording(session.recording) || {}),
      status: segments.length ? "speaker timing ready" : "transcribed",
      durationSeconds: Number(payload.durationSeconds || session.recording?.durationSeconds || 0),
      fileName: payload.fileName || file.name || "session-audio",
    };
    audioUploadStatus = "complete";
    speakerTimingStatus = "complete";
    elements.audioUploadFile.value = "";
    elements.audioUploadMessage.textContent = segments.length
      ? "Speaker segments are ready. Match each AI speaker label to a player or DM, then apply the reviewed timing."
      : "Transcript came back, but no speaker segments were found. Manual speaking time is still available.";
    render();
  } catch (error) {
    speakerTimingStatus = "error";
    elements.audioUploadMessage.textContent =
      error.message === "Failed to fetch"
        ? `Could not reach the AI backend at ${API_BASE_URL}.`
        : error.message || "Speaker timing failed.";
    renderSpeakingTime();
    saveState();
  }
});

elements.saveSpeakerStats.addEventListener("click", () => {
  const session = getCurrentSession();
  if (!session) return;
  const inputs = elements.speakerList.querySelectorAll("[data-speaker-id]");
  session.speakerStats = Array.from(inputs).map((input) => {
    const speaker = getSpeakers().find(
      (entry) => entry.id === input.dataset.speakerId && entry.type === input.dataset.speakerType,
    );
    return {
      speakerId: input.dataset.speakerId,
      speakerType: input.dataset.speakerType,
      name: speaker?.name || "Unknown speaker",
      minutes: Math.max(0, Number(input.value || 0)),
    };
  });
  elements.speakerStatsMessage.textContent = "Speaking time saved for this session.";
  render();
});

function deleteSpeakingStat(speakerId, speakerType, scope) {
  if (!speakerId || !speakerType) return;
  if (scope === "campaign") {
    const activeCampaign = getActiveCampaign();
    getCampaignSessions(state, activeCampaign?.id).forEach((session) => {
      session.speakerStats = (session.speakerStats || []).filter(
        (stat) => !(stat.speakerId === speakerId && stat.speakerType === speakerType),
      );
    });
    elements.speakerStatsMessage.textContent = "Speaker removed from active campaign totals.";
    render();
    return;
  }

  const session = getCurrentSession();
  if (!session) return;
  session.speakerStats = (session.speakerStats || []).filter(
    (stat) => !(stat.speakerId === speakerId && stat.speakerType === speakerType),
  );
  elements.speakerStatsMessage.textContent = "Speaker removed from this session chart.";
  render();
}

elements.sessionSpeakingChart.addEventListener("click", (event) => {
  const button = event.target.closest("[data-delete-speaking-stat]");
  if (!button) return;
  deleteSpeakingStat(button.dataset.speakerId, button.dataset.speakerType, "session");
});

elements.campaignSpeakingChart.addEventListener("click", (event) => {
  const button = event.target.closest("[data-delete-speaking-stat]");
  if (!button) return;
  deleteSpeakingStat(button.dataset.speakerId, button.dataset.speakerType, "campaign");
});

elements.applySpeakerReview.addEventListener("click", () => {
  const session = getCurrentSession();
  if (!session) return;

  const mappings = new Map();
  elements.speakerMapList.querySelectorAll("[data-speaker-map-label]").forEach((select) => {
    mappings.set(select.dataset.speakerMapLabel, select.value);
  });

  const speakers = getSpeakers();
  const totals = new Map();
  session.speakerSegments = (session.speakerSegments || []).map((segment) => {
    if (segment.source !== "ai") return segment;
    const mappingValue = mappings.get(segment.speakerLabel || "Unknown speaker");
    if (!mappingValue) return { ...segment, assignedSpeakerId: "", assignedSpeakerType: "", assignedSpeakerName: "" };

    const [speakerType, speakerId] = mappingValue.split(":");
    const speaker = speakers.find((entry) => entry.type === speakerType && entry.id === speakerId);
    if (!speaker) return segment;

    const minutes = Math.max(0, Number(segment.minutes || 0));
    const key = `${speaker.type}:${speaker.id}`;
    const current = totals.get(key) || {
      speakerId: speaker.id,
      speakerType: speaker.type,
      name: speaker.name,
      minutes: 0,
    };
    current.minutes += minutes;
    totals.set(key, current);

    return {
      ...segment,
      assignedSpeakerId: speaker.id,
      assignedSpeakerType: speaker.type,
      assignedSpeakerName: speaker.name,
    };
  });

  session.speakerStats = Array.from(totals.values()).map((entry) => ({
    ...entry,
    minutes: Math.round(entry.minutes * 10) / 10,
  }));
  session.speakerReviewStatus = session.speakerStats.length ? "reviewed" : "needs-review";
  elements.speakerStatsMessage.textContent = session.speakerStats.length
    ? "Reviewed speaker timing applied to the charts."
    : "Choose at least one player or DM before applying speaker timing.";
  elements.audioUploadMessage.textContent = session.speakerStats.length
    ? "Reviewed speaker timing was applied to the speaking charts."
    : "Choose at least one player or DM before applying speaker timing.";
  render();
});

elements.speakerSegmentForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const session = getCurrentSession();
  if (!session) return;
  const speakerLabel = elements.segmentSpeakerLabel.value.trim();
  const note = elements.segmentNote.value.trim();
  const minutes = Math.max(0, Number(elements.segmentMinutes.value || 0));
  if (!speakerLabel && !note && !minutes) return;
  session.speakerSegments = [
    ...(session.speakerSegments || []),
    {
      id: createId("segment"),
      speakerLabel: speakerLabel || "Unknown speaker",
      minutes,
      note,
      source: "manual",
      createdAt: new Date().toISOString(),
    },
  ];
  session.speakerReviewStatus = "manual";
  elements.speakerSegmentForm.reset();
  elements.segmentMinutes.value = 0;
  render();
});

elements.speakerSegmentList.addEventListener("click", (event) => {
  const deleteId = event.target.dataset.deleteSegment;
  const session = getCurrentSession();
  if (!deleteId || !session) return;
  session.speakerSegments = (session.speakerSegments || []).filter((segment) => segment.id !== deleteId);
  render();
});

elements.transcriptField.addEventListener("input", () => {
  updateGenerateRecapButton();
  renderAiSetupStatus(getCurrentSession());
});

elements.saveApiBaseUrl.addEventListener("click", () => {
  const nextUrl = elements.apiBaseUrl.value.trim().replace(/\/+$/, "");
  if (!nextUrl) return;
  API_BASE_URL = nextUrl;
  localStorage.setItem("dnd-club-api-base-url", API_BASE_URL);
  state.settings.aiSetup = {
    ...clone(initialState.settings.aiSetup),
    message: "Backend URL saved. Run Check AI setup to verify it.",
  };
  const session = getCurrentSession();
  if (session && session.aiStatus === "error") {
    session.aiStatus = elements.transcriptField.value.trim() ? "ready" : "not-ready";
  }
  render();
});

elements.clearApiBaseUrl.addEventListener("click", () => {
  localStorage.removeItem("dnd-club-api-base-url");
  API_BASE_URL = ["localhost", "127.0.0.1"].includes(window.location.hostname)
    ? "http://localhost:3001"
    : "https://dnd-club-session-organizer.onrender.com";
  state.settings.aiSetup = {
    ...clone(initialState.settings.aiSetup),
    message: "Default backend URL restored. Run Check AI setup to verify it.",
  };
  render();
});

elements.checkAiSetup.addEventListener("click", async () => {
  state.settings.aiSetup = {
    ...(state.settings.aiSetup || clone(initialState.settings.aiSetup)),
    message: "Checking AI setup...",
  };
  renderProgressNotes();

  try {
    const response = await fetch(`${API_BASE_URL}/api/health`);
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || "Backend health check failed.");
    }
    state.settings.aiSetup = {
      lastCheckedAt: new Date().toISOString(),
      backendReachable: true,
      hasOpenAIKey: Boolean(payload.hasOpenAIKey),
      service: payload.service || "D&D Club AI backend",
      message: payload.hasOpenAIKey
        ? "Backend and OpenAI are connected."
        : "Backend is reachable, but the OpenAI API key is missing on Render.",
    };
  } catch (error) {
    state.settings.aiSetup = {
      lastCheckedAt: new Date().toISOString(),
      backendReachable: false,
      hasOpenAIKey: null,
      service: "",
      message:
        error.message === "Failed to fetch"
          ? "Could not reach the backend. Check the Backend URL or wake up the Render service."
          : error.message || "Could not check AI setup.",
    };
  }
  render();
});

elements.generateRecap.addEventListener("click", async () => {
  const session = getCurrentSession();
  if (!session) return;
  if (state.settings.manualMode) {
    elements.aiToolsNote.textContent = "Manual Mode is on. Turn it off to use AI recap generation.";
    return;
  }
  const transcript = elements.transcriptField.value.trim();
  if (!transcript) {
    updateGenerateRecapButton();
    return;
  }

  const activeCampaign = getActiveCampaign();
  const sessionLabel = getSessionLabel(session, getSessionIndex(session));
  const existingRecap = getCurrentRecapDraft();
  session.aiStatus = "processing";
  elements.aiStatus.textContent = formatStatus(session.aiStatus);
  elements.aiToolsNote.textContent = getAiToolsMessage(session);
  updateGenerateRecapButton();

  try {
    const response = await fetch(`${API_BASE_URL}/api/generate-recap`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        transcript,
        campaignName: activeCampaign?.name || "Active campaign",
        sessionLabel,
        existingRecap,
      }),
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || "AI recap generation failed.");
    }

    elements.recapSummaryField.value = payload.summary || "";
    elements.recapEventsField.value = payload.events || "";
    elements.recapNpcsField.value = payload.npcs || "";
    elements.recapLocationsField.value = payload.locations || "";
    elements.recapQuestsField.value = payload.quests || "";
    elements.recapThreadsField.value = payload.unresolvedThreads || "";
    session.aiStatus = "complete";
    elements.aiStatus.textContent = formatStatus(session.aiStatus);
    elements.aiToolsNote.textContent = getAiToolsMessage(session);
  } catch (error) {
    session.aiStatus = "error";
    elements.aiStatus.textContent = formatStatus(session.aiStatus);
    elements.aiToolsNote.textContent =
      error.message === "Failed to fetch"
        ? `Could not reach the AI backend at ${API_BASE_URL}. Start the backend locally or deploy it on Render.`
        : error.message || getAiToolsMessage(session);
  } finally {
    saveState();
    updateGenerateRecapButton();
  }
});

elements.progressForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const session = getCurrentSession();
  if (!session) return;
  session.transcript = elements.transcriptField.value.trim();
  session.recap = {
    summary: elements.recapSummaryField.value.trim(),
    events: elements.recapEventsField.value.trim(),
    npcs: elements.recapNpcsField.value.trim(),
    locations: elements.recapLocationsField.value.trim(),
    quests: elements.recapQuestsField.value.trim(),
    unresolvedThreads: elements.recapThreadsField.value.trim(),
  };
  render();
});

elements.exportSessionMarkdown.addEventListener("click", () => {
  const session = getCurrentSession();
  if (!session) return;
  const fileName = `${slugifyFileName(getSessionLabel(session, getSessionIndex(session)))}.md`;
  downloadTextFile(fileName, buildSessionMarkdown(session));
});

elements.exportCampaignMarkdown.addEventListener("click", () => {
  const campaign = getActiveCampaign();
  if (!campaign) return;
  downloadTextFile(`${slugifyFileName(campaign.name)}-campaign-story.md`, buildCampaignMarkdown(campaign));
});

render();
