import React, { useEffect, useMemo, useRef, useState } from "react";

const STORAGE_KEY = "vms_records_v2";
const AUDIT_KEY = "vms_audit_v2";
const NOTIFY_ENDPOINT = window.NOTIFY_ENDPOINT || "/api/notify-host";
const DEFAULT_COMPANY = "The Church of Jesus Christ of Latter-Day Saints";

const EMPLOYEES = [
  {
    id: "E1001",
    name: "Alexandra Reed",
    email: "alexandra.reed@hq.example",
    phone: "+1-555-0100",
    location: "HQ Building A / Floor 4 / Suite 402",
    department: "Operations",
    title: "Director of Operations",
    manager: "Riley Carter",
    status: "Active",
  },
  {
    id: "E1002",
    name: "Marcus Alvarez",
    email: "marcus.alvarez@hq.example",
    phone: "+1-555-0114",
    location: "HQ Building A / Floor 2 / Room 2-118",
    department: "Security",
    title: "Head of Security",
    manager: "Riley Carter",
    status: "Active",
  },
  {
    id: "E1003",
    name: "Priya Shah",
    email: "priya.shah@hq.example",
    phone: "+1-555-0135",
    location: "HQ Building A / Floor 5 / Suite 503",
    department: "Finance",
    title: "Finance Controller",
    manager: "Noah Kim",
    status: "Active",
  },
  {
    id: "E1004",
    name: "Dana Mitchell",
    email: "dana.mitchell@hq.example",
    phone: "+1-555-0148",
    location: "HQ Building A / Floor 3 / Suite 314",
    department: "HR",
    title: "HR Business Partner",
    manager: "Noah Kim",
    status: "Inactive",
  },
  {
    id: "E1005",
    name: "Ethan Brooks",
    email: "ethan.brooks@hq.example",
    phone: "+1-555-0161",
    location: "HQ Building B / Floor 1 / Suite 114",
    department: "Engineering",
    title: "Engineering Manager",
    manager: "Riley Carter",
    status: "Active",
  },
  {
    id: "E1006",
    name: "Sophia Chen",
    email: "sophia.chen@hq.example",
    phone: "+1-555-0172",
    location: "HQ Building B / Floor 3 / Suite 306",
    department: "Product",
    title: "Senior Product Manager",
    manager: "Riley Carter",
    status: "Active",
  },
  {
    id: "E1007",
    name: "Liam Patel",
    email: "liam.patel@hq.example",
    phone: "+1-555-0180",
    location: "HQ Building A / Floor 6 / Suite 611",
    department: "Legal",
    title: "Corporate Counsel",
    manager: "Noah Kim",
    status: "Active",
  },
  {
    id: "E1008",
    name: "Mia Johnson",
    email: "mia.johnson@hq.example",
    phone: "+1-555-0194",
    location: "HQ Building C / Floor 2 / Suite 209",
    department: "Marketing",
    title: "Marketing Director",
    manager: "Riley Carter",
    status: "Active",
  },
  {
    id: "E1009",
    name: "Noah Turner",
    email: "noah.turner@hq.example",
    phone: "+1-555-0201",
    location: "HQ Building C / Floor 4 / Suite 417",
    department: "Sales",
    title: "Enterprise Account Executive",
    manager: "Riley Carter",
    status: "Active",
  },
  {
    id: "E1010",
    name: "Grace Walker",
    email: "grace.walker@hq.example",
    phone: "+1-555-0216",
    location: "HQ Building A / Floor 2 / Suite 224",
    department: "IT",
    title: "IT Operations Lead",
    manager: "Noah Kim",
    status: "Active",
  },
  {
    id: "E1011",
    name: "Benjamin Flores",
    email: "benjamin.flores@hq.example",
    phone: "+1-555-0228",
    location: "HQ Building B / Floor 5 / Suite 503",
    department: "Procurement",
    title: "Procurement Manager",
    manager: "Noah Kim",
    status: "Active",
  },
  {
    id: "E1012",
    name: "Ava Robinson",
    email: "ava.robinson@hq.example",
    phone: "+1-555-0239",
    location: "HQ Building C / Floor 1 / Suite 108",
    department: "Customer Success",
    title: "Customer Success Director",
    manager: "Riley Carter",
    status: "Active",
  },
  {
    id: "E1013",
    name: "Lucas Garcia",
    email: "lucas.garcia@hq.example",
    phone: "+1-555-0245",
    location: "HQ Building A / Floor 7 / Suite 705",
    department: "Executive",
    title: "Chief Technology Officer",
    manager: "Board of Directors",
    status: "Active",
  },
];

const WATCHLIST = [
  { name: "john doe", risk: "High", notes: "Internal POI list" },
  { name: "jane offender", risk: "Critical", notes: "No-fly style alert" },
];

const VISIT_REASONS = [
  "Interview",
  "Client Meeting",
  "Vendor Delivery",
  "Maintenance / Repair",
  "Training Session",
  "Board Meeting",
  "Security Audit",
  "Contractor Visit",
  "IT Support",
  "Facility Tour",
];
const OTHER_REASON_VALUE = "Other";

const CHECK_IN_METHODS = [
  "QR",
  "PIN",
  "OTP",
  "Manual Lookup",
  "Facial (Optional)",
];
const KIOSKS = ["Lobby-K1", "Lobby-K2", "Reception-Desk"];

const EMPTY_PRE_REG = {
  visitorFirstName: "",
  visitorLastName: "",
  visitorEmail: "",
  visitorPhone: "",
  company: DEFAULT_COMPANY,
  hostId: "",
  visitDate: null,
  visitTime: null,
  expectedDurationMinutes: "60",
  reason: "",
  reasonCustom: "",
  accessLevel: "Standard",
  parking: "",
  ndaRequired: false,
  accommodations: "",
};

const EMPTY_CHECKIN = {
  lookupValue: "",
  method: "QR",
  kioskId: "Lobby-K1",
  company: DEFAULT_COMPANY,
  verifyEmail: "",
  verifyPhone: "",
  verifyReason: "",
  verifyHostId: "",
  verifyDate: "",
  verifyTime: "",
  verifyDurationMinutes: "",
  verifyAccessLevel: "",
  verifyParking: "",
  verifyAccommodations: "",
  signature: "",
};

const EMPTY_SIGNOUT = {
  lookupValue: "",
  method: "QR",
  kioskId: "Lobby-K1",
};

const LANGUAGE_PACK = {
  English: {
    eyebrow: "Enterprise Visitor Management System",
    title: "Secure Guest Management Console",
    role: "Role",
    language: "Language",
    admin: "Admin",
    frontDesk: "Front Desk",
    sectionPreReg: "Guest Sign-In",
    sectionCheckIn: "Pre-Registered Guest",
    sectionSignOut: "Guest Sign-Out",
    guestSignInTab: "Guest Sign In",
    returningGuestTab: "Pre-Registered Guest",
    guestSignOutTab: "Guest Sign Out",
    sectionHost: "Host Coordination & Credentials",
    sectionOccupancy: "Occupancy, Emergency, and Compliance",
    sectionAnalytics: "Reporting & Analytics",
    sectionRecords: "Visitor Records",
    createPreReg: "Complete Guest Sign In",
    visitorFirstName: "Visitor First Name *",
    visitorLastName: "Visitor Last Name *",
    visitorEmail: "Visitor Email *",
    visitorPhone: "Visitor Phone",
    company: "Company",
    reasonForVisit: "Reason For Visit *",
    host: "Host *",
    date: "Date *",
    time: "Time *",
    expectedDuration: "Expected Duration (minutes)",
    accessLevel: "Access Level",
    parking: "Parking Reservation",
    ndaRequired: "NDA Required",
    specialAccommodations: "Special Accommodations",
    checkInMethod: "Method",
    kioskId: "Kiosk ID",
    lookupValue: "Lookup Value (QR / PIN / OTP / Email / Visit ID)",
    verifyEmail: "Verify Email",
    verifyPhone: "Verify Phone",
    verifyReason: "Verify Reason",
    confirmEmail: "Confirm Email",
    confirmPhone: "Confirm Phone",
    confirmReason: "Confirm Reason",
    confirmHost: "Confirm Host",
    confirmDate: "Confirm Date",
    confirmTime: "Confirm Time",
    confirmDuration: "Confirm Expected Duration (minutes)",
    confirmAccessLevel: "Confirm Access Level",
    confirmParking: "Confirm Parking Reservation",
    confirmAccommodations: "Confirm Special Accommodations",
    ndaSignature: "NDA Signature (if required)",
    completeCheckIn: "Complete Returning Guest Check-In",
    completeSignOut: "Complete Guest Sign-Out",
    runAutoExpiry: "Run Auto Expiry",
    signOutMethod: "Sign-Out Method",
    signOutLookupValue: "Lookup Value (QR / PIN / OTP / Email / Visit ID)",
    signOutSummary: "Active Visit Ready for Sign-Out",
    signOutStatusLabel: "Current Status",
    signOutLocationLabel: "Checked In At",
    signOutTimeLabel: "Checked In Time",
    signOutEmpty: "Find an active visit to sign the guest out.",
    signOutActiveList: "Checked-In Guests",
    signOutActiveEmpty: "No guests are currently checked in.",
    signOutSelect: "Select For Sign-Out",
    signOutDone: "Guest sign-out complete.",
    preRegDone: "Pre-registration complete.",
    visitId: "Visit ID",
    qrToken: "QR Token",
    pin: "PIN",
    missingPrefix: "Please complete the following required fields",
    invalidEmail: "Enter a valid visitor email.",
    photoRequired: "Capture a guest photo before submitting sign-in.",
    cameraTitle: "Guest Photo Capture (Required)",
    cameraHint: "Take the photo here before completing guest sign in.",
  },
  Spanish: {
    eyebrow: "Sistema Empresarial de Gestion de Visitantes",
    title: "Consola Segura de Gestion de Invitados",
    role: "Rol",
    language: "Idioma",
    admin: "Administrador",
    frontDesk: "Recepcion",
    sectionPreReg: "Registro de Invitados",
    sectionCheckIn: "Invitado Pre-Registrado",
    sectionSignOut: "Salida de Invitados",
    guestSignInTab: "Registro de Invitados",
    returningGuestTab: "Invitado Pre-Registrado",
    guestSignOutTab: "Salida de Invitados",
    sectionHost: "Coordinacion con Anfitrion y Credenciales",
    sectionOccupancy: "Ocupacion, Emergencia y Cumplimiento",
    sectionAnalytics: "Reportes y Analitica",
    sectionRecords: "Registros de Visitantes",
    createPreReg: "Completar Registro de Invitado",
    visitorFirstName: "Nombre del Visitante *",
    visitorLastName: "Apellido del Visitante *",
    visitorEmail: "Correo del Visitante *",
    visitorPhone: "Telefono del Visitante",
    company: "Empresa",
    reasonForVisit: "Motivo de la Visita *",
    host: "Anfitrion *",
    date: "Fecha *",
    time: "Hora *",
    expectedDuration: "Duracion Esperada (minutos)",
    accessLevel: "Nivel de Acceso",
    parking: "Reserva de Estacionamiento",
    ndaRequired: "NDA Requerido",
    specialAccommodations: "Acomodaciones Especiales",
    checkInMethod: "Metodo",
    kioskId: "ID del Kiosco",
    lookupValue: "Valor de Busqueda (QR / PIN / OTP / Correo / ID de Visita)",
    verifyEmail: "Verificar Correo",
    verifyPhone: "Verificar Telefono",
    verifyReason: "Verificar Motivo",
    confirmEmail: "Confirmar Correo",
    confirmPhone: "Confirmar Telefono",
    confirmReason: "Confirmar Motivo",
    confirmHost: "Confirmar Anfitrion",
    confirmDate: "Confirmar Fecha",
    confirmTime: "Confirmar Hora",
    confirmDuration: "Confirmar Duracion Esperada (minutos)",
    confirmAccessLevel: "Confirmar Nivel de Acceso",
    confirmParking: "Confirmar Reserva de Estacionamiento",
    confirmAccommodations: "Confirmar Acomodaciones Especiales",
    ndaSignature: "Firma NDA (si es requerido)",
    completeCheckIn: "Completar Registro",
    completeSignOut: "Completar Salida del Invitado",
    runAutoExpiry: "Ejecutar Expiracion Automatica",
    signOutMethod: "Metodo de Salida",
    signOutLookupValue:
      "Valor de Busqueda (QR / PIN / OTP / Correo / ID de Visita)",
    signOutSummary: "Visita Activa Lista para Salida",
    signOutStatusLabel: "Estado Actual",
    signOutLocationLabel: "Registro en",
    signOutTimeLabel: "Hora de Entrada",
    signOutEmpty: "Busque una visita activa para registrar la salida del invitado.",
    signOutActiveList: "Invitados Registrados",
    signOutActiveEmpty: "No hay invitados registrados en este momento.",
    signOutSelect: "Seleccionar para Salida",
    signOutDone: "Salida del invitado completada.",
    preRegDone: "Pre-registro completado.",
    visitId: "ID de Visita",
    qrToken: "Token QR",
    pin: "PIN",
    missingPrefix: "Complete los siguientes campos obligatorios",
    invalidEmail: "Ingrese un correo electronico valido del visitante.",
    photoRequired: "Capture una foto del invitado antes de enviar el registro.",
    cameraTitle: "Captura de Foto del Invitado (Requerida)",
    cameraHint: "Tome la foto aqui antes de completar el registro.",
  },
  French: {
    eyebrow: "Systeme d'Accueil des Visiteurs Entreprise",
    title: "Console Securisee de Gestion des Visiteurs",
    role: "Role",
    language: "Langue",
    admin: "Administrateur",
    frontDesk: "Accueil",
    sectionPreReg: "Enregistrement des Invites",
    sectionCheckIn: "Visiteur Preinscrit",
    sectionSignOut: "Sortie des Invites",
    guestSignInTab: "Enregistrement Invite",
    returningGuestTab: "Visiteur Preinscrit",
    guestSignOutTab: "Sortie des Invites",
    sectionHost: "Coordination Hote et Badges",
    sectionOccupancy: "Occupation, Urgence et Conformite",
    sectionAnalytics: "Rapports et Analytique",
    sectionRecords: "Registres des Visiteurs",
    createPreReg: "Terminer l'Enregistrement Invite",
    visitorFirstName: "Prenom du Visiteur *",
    visitorLastName: "Nom du Visiteur *",
    visitorEmail: "E-mail du Visiteur *",
    visitorPhone: "Telephone du Visiteur",
    company: "Entreprise",
    reasonForVisit: "Motif de la Visite *",
    host: "Hote *",
    date: "Date *",
    time: "Heure *",
    expectedDuration: "Duree Prevue (minutes)",
    accessLevel: "Niveau d'Acces",
    parking: "Reservation de Parking",
    ndaRequired: "NDA Requis",
    specialAccommodations: "Amenagements Speciaux",
    checkInMethod: "Methode",
    kioskId: "ID du Kiosque",
    lookupValue: "Valeur de Recherche (QR / PIN / OTP / E-mail / ID Visite)",
    verifyEmail: "Verifier l'E-mail",
    verifyPhone: "Verifier le Telephone",
    verifyReason: "Verifier le Motif",
    confirmEmail: "Confirmer l'E-mail",
    confirmPhone: "Confirmer le Telephone",
    confirmReason: "Confirmer le Motif",
    confirmHost: "Confirmer l'Hote",
    confirmDate: "Confirmer la Date",
    confirmTime: "Confirmer l'Heure",
    confirmDuration: "Confirmer la Duree Prevue (minutes)",
    confirmAccessLevel: "Confirmer le Niveau d'Acces",
    confirmParking: "Confirmer la Reservation de Parking",
    confirmAccommodations: "Confirmer les Amenagements Speciaux",
    ndaSignature: "Signature NDA (si requise)",
    completeCheckIn: "Terminer l'Enregistrement",
    completeSignOut: "Terminer la Sortie de l'Invite",
    runAutoExpiry: "Executer l'Expiration Auto",
    signOutMethod: "Methode de Sortie",
    signOutLookupValue:
      "Valeur de Recherche (QR / PIN / OTP / E-mail / ID Visite)",
    signOutSummary: "Visite Active Prete pour la Sortie",
    signOutStatusLabel: "Statut Actuel",
    signOutLocationLabel: "Enregistre a",
    signOutTimeLabel: "Heure d'Entree",
    signOutEmpty: "Recherchez une visite active pour terminer la sortie de l'invite.",
    signOutActiveList: "Invites Enregistres",
    signOutActiveEmpty: "Aucun invite n'est actuellement enregistre.",
    signOutSelect: "Selectionner pour la Sortie",
    signOutDone: "Sortie de l'invite terminee.",
    preRegDone: "Pre-inscription terminee.",
    visitId: "ID de Visite",
    qrToken: "Jeton QR",
    pin: "PIN",
    missingPrefix: "Veuillez remplir les champs obligatoires suivants",
    invalidEmail: "Saisissez une adresse e-mail visiteur valide.",
    photoRequired:
      "Capturez une photo de l'invite avant de soumettre l'enregistrement.",
    cameraTitle: "Capture Photo Invite (Obligatoire)",
    cameraHint: "Prenez la photo ici avant de terminer l'enregistrement.",
  },
};

const REQUIRED_PRE_REG_FIELDS = [
  "visitorFirstName",
  "visitorLastName",
  "visitorEmail",
  "hostId",
  "visitDate",
  "visitTime",
  "reason",
];

const REQUIRED_FIELD_LABELS = {
  English: {
    visitorFirstName: "Visitor First Name",
    visitorLastName: "Visitor Last Name",
    visitorEmail: "Visitor Email",
    hostId: "Host",
    visitDate: "Date",
    visitTime: "Time",
    reason: "Reason For Visit",
  },
  Spanish: {
    visitorFirstName: "Nombre del Visitante",
    visitorLastName: "Apellido del Visitante",
    visitorEmail: "Correo del Visitante",
    hostId: "Anfitrion",
    visitDate: "Fecha",
    visitTime: "Hora",
    reason: "Motivo de la Visita",
  },
  French: {
    visitorFirstName: "Prenom du Visiteur",
    visitorLastName: "Nom du Visiteur",
    visitorEmail: "E-mail du Visiteur",
    hostId: "Hote",
    visitDate: "Date",
    visitTime: "Heure",
    reason: "Motif de la Visite",
  },
};

function loadJson(key, fallback) {
  try {
    return JSON.parse(localStorage.getItem(key)) || fallback;
  } catch {
    return fallback;
  }
}

function saveJson(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function fmtDateTime(iso) {
  if (!iso) return "-";
  return new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(iso));
}

function fmtTime(iso) {
  if (!iso) return "-";
  return new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(iso));
}

function dayKey(date) {
  return new Date(date).toISOString().slice(0, 10);
}

function createVisitId() {
  return `VMS-${new Date().toISOString().slice(0, 10).replaceAll("-", "")}-${Math.floor(
    1000 + Math.random() * 9000,
  )}`;
}

function createPin() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function createToken() {
  return crypto.randomUUID().replaceAll("-", "").slice(0, 20);
}

function buildVisitorFullName({ visitorFirstName, visitorLastName }) {
  return `${visitorFirstName} ${visitorLastName}`.replace(/\s+/g, " ").trim();
}

function formatPhone(value) {
  const digits = value.replace(/\D/g, "").slice(0, 10);
  if (!digits) return "";
  if (digits.length < 4) return `(${digits}`;
  if (digits.length < 7) {
    return `(${digits.slice(0, 3)}) ${digits.slice(3)}`;
  }
  return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
}

function getNowForDateInput() {
  const now = new Date();
  const local = new Date(now.getTime() - now.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
}

function getNowForTimeInput() {
  const now = new Date();
  return now.toTimeString().slice(0, 5);
}

function getDefaultPreReg() {
  return {
    ...EMPTY_PRE_REG,
    visitDate: getNowForDateInput(),
    visitTime: getNowForTimeInput(),
  };
}

function checkWatchlist(fullName) {
  const normalized = fullName.toLowerCase().trim().replace(/\s+/g, " ");
  const hit = WATCHLIST.find((entry) => entry.name === normalized);
  if (!hit) {
    return { status: "clear", label: "Clear", detail: "No watchlist match." };
  }
  return {
    status: "review",
    label: "Security Hold",
    detail: `Watchlist match (${hit.risk}). ${hit.notes}`,
  };
}

function App() {
  const [role, setRole] = useState("Front Desk");
  const [language, setLanguage] = useState("English");
  const [activeIntakeTab, setActiveIntakeTab] = useState("guestSignIn");
  const [theme, setTheme] = useState(
    () => localStorage.getItem("vms_theme") || "dark",
  );
  const [clock, setClock] = useState(() =>
    fmtDateTime(new Date().toISOString()),
  );
  const [preReg, setPreReg] = useState(() => getDefaultPreReg());
  const [checkIn, setCheckIn] = useState(EMPTY_CHECKIN);
  const [signOut, setSignOut] = useState(EMPTY_SIGNOUT);
  const [records, setRecords] = useState(() => loadJson(STORAGE_KEY, []));
  const [audit, setAudit] = useState(() => loadJson(AUDIT_KEY, []));
  const [activeVisitId, setActiveVisitId] = useState(null);
  const [emergencyMode, setEmergencyMode] = useState(false);
  const [retentionDays, setRetentionDays] = useState(90);
  const [notificationChannels, setNotificationChannels] = useState({
    teams: true,
    email: true,
    sms: false,
    push: false,
  });
  const [screeningResult, setScreeningResult] = useState({
    status: "waiting",
    label: "Awaiting Check-In",
    detail: "No screening executed yet.",
  });

  const [photoDataUrl, setPhotoDataUrl] = useState("");
  const [cameraReady, setCameraReady] = useState(false);
  const [qrScanning, setQrScanning] = useState(false);
  const [qrScanError, setQrScanError] = useState("");
  const videoRef = useRef(null);
  const canvasRef = useRef(null);
  const streamRef = useRef(null);
  const qrVideoRef = useRef(null);
  const qrStreamRef = useRef(null);
  const qrRafRef = useRef(null);
  const lastCompanyAutofillIdRef = useRef(null);

  useEffect(() => {
    const timer = setInterval(
      () => setClock(fmtDateTime(new Date().toISOString())),
      30000,
    );
    return () => clearInterval(timer);
  }, []);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem("vms_theme", theme);
  }, [theme]);

  useEffect(() => {
    saveJson(STORAGE_KEY, records);
  }, [records]);

  useEffect(() => {
    saveJson(AUDIT_KEY, audit);
  }, [audit]);

  useEffect(() => {
    return () => {
      stopCamera();
      stopQrScanner();
    };
  }, []);

  useEffect(() => {
    const qrEnabled =
      (activeIntakeTab === "returningGuest" && checkIn.method === "QR") ||
      (activeIntakeTab === "guestSignOut" && signOut.method === "QR");
    if (!qrEnabled) {
      stopQrScanner();
    }
  }, [activeIntakeTab, checkIn.method, signOut.method]);

  const employees = useMemo(
    () => EMPLOYEES.filter((emp) => emp.status === "Active"),
    [],
  );

  const selectedHost = useMemo(
    () => employees.find((emp) => emp.id === preReg.hostId) || null,
    [employees, preReg.hostId],
  );

  const activeRecord = useMemo(
    () => records.find((entry) => entry.id === activeVisitId) || null,
    [records, activeVisitId],
  );

  const occupancy = useMemo(
    () =>
      records.filter(
        (entry) => entry.status === "checked_in" && !entry.evacuatedAt,
      ),
    [records],
  );

  const todaysRecords = useMemo(() => {
    const today = dayKey(new Date());
    return records.filter((entry) => dayKey(entry.createdAt) === today);
  }, [records]);

  const metrics = useMemo(() => {
    const now = Date.now();
    const sevenDaysAgo = now - 7 * 24 * 60 * 60 * 1000;
    const monthAgo = now - 30 * 24 * 60 * 60 * 1000;

    const dailyCount = records.filter(
      (entry) => dayKey(entry.createdAt) === dayKey(new Date()),
    ).length;
    const weeklyCount = records.filter(
      (entry) => new Date(entry.createdAt).getTime() >= sevenDaysAgo,
    ).length;
    const monthlyCount = records.filter(
      (entry) => new Date(entry.createdAt).getTime() >= monthAgo,
    ).length;

    const noShow = records.filter(
      (entry) =>
        entry.status === "pre_registered" && new Date(entry.endAt) < new Date(),
    ).length;
    const repeat = Object.values(
      records.reduce((acc, entry) => {
        acc[entry.visitorEmail] = (acc[entry.visitorEmail] || 0) + 1;
        return acc;
      }, {}),
    ).filter((count) => count > 1).length;

    const completed = records.filter(
      (entry) => entry.checkedInAt && entry.checkedOutAt,
    );
    const avgDuration = completed.length
      ? Math.round(
          completed.reduce(
            (sum, entry) =>
              sum +
              (new Date(entry.checkedOutAt) - new Date(entry.checkedInAt)),
            0,
          ) /
            completed.length /
            60000,
        )
      : 0;

    return {
      dailyCount,
      weeklyCount,
      monthlyCount,
      noShow,
      repeat,
      avgDuration,
    };
  }, [records]);

  const byReason = useMemo(() => {
    const grouped = {};
    records.forEach((entry) => {
      grouped[entry.reason] = (grouped[entry.reason] || 0) + 1;
    });
    return Object.entries(grouped).sort((a, b) => b[1] - a[1]);
  }, [records]);

  const byHour = useMemo(() => {
    const buckets = Array.from({ length: 24 }, (_, hour) => ({
      hour,
      count: 0,
    }));
    records.forEach((entry) => {
      const hour = new Date(entry.createdAt).getHours();
      buckets[hour].count += 1;
    });
    return buckets;
  }, [records]);

  const t = (key) =>
    LANGUAGE_PACK[language]?.[key] || LANGUAGE_PACK.English[key] || key;
  const isAdmin = role === "Admin";

  function logAudit(action, details) {
    setAudit((prev) =>
      [
        {
          id: crypto.randomUUID(),
          at: new Date().toISOString(),
          action,
          details,
        },
        ...prev,
      ].slice(0, 200),
    );
  }

  function updatePreReg(key, value) {
    setPreReg((prev) => ({ ...prev, [key]: value }));
  }

  function updateCheckIn(key, value) {
    setCheckIn((prev) => ({ ...prev, [key]: value }));
  }

  function updateSignOut(key, value) {
    setSignOut((prev) => ({ ...prev, [key]: value }));
  }

  async function startCamera() {
    if (streamRef.current) return;
    if (qrStreamRef.current) {
      stopQrScanner();
    }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: true,
        audio: false,
      });
      streamRef.current = stream;
      videoRef.current.srcObject = stream;
      setCameraReady(true);
      logAudit("camera.start", "Camera stream started");
    } catch {
      alert("Camera access blocked. Allow camera permissions.");
    }
  }

  function stopCamera() {
    const stream = streamRef.current;
    if (!stream) return;
    stream.getTracks().forEach((track) => track.stop());
    streamRef.current = null;
    setCameraReady(false);
  }

  function capturePhoto() {
    if (!streamRef.current) {
      alert("Start camera before capture.");
      return;
    }
    const canvas = canvasRef.current;
    const video = videoRef.current;
    if (!canvas || !video) {
      return;
    }

    const sourceWidth = video.videoWidth || 0;
    const sourceHeight = video.videoHeight || 0;
    if (!sourceWidth || !sourceHeight) {
      alert("Camera is still warming up. Try capture again.");
      return;
    }

    const squareSize = Math.min(sourceWidth, sourceHeight);
    const sx = (sourceWidth - squareSize) / 2;
    const sy = (sourceHeight - squareSize) / 2;
    canvas.width = squareSize;
    canvas.height = squareSize;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(
      video,
      sx,
      sy,
      squareSize,
      squareSize,
      0,
      0,
      squareSize,
      squareSize,
    );
    setPhotoDataUrl(canvas.toDataURL("image/jpeg", 0.9));
    logAudit("visitor.photo.captured", "Photo captured at kiosk");
  }

  function clearPhoto() {
    setPhotoDataUrl("");
  }

  async function startQrScanner() {
    if (qrStreamRef.current) return;
    setQrScanError("");

    if (!("BarcodeDetector" in window)) {
      setQrScanError(
        "QR scanning is not supported in this browser. Enter the code manually.",
      );
      return;
    }

    if (streamRef.current) {
      stopCamera();
    }

    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: "environment" } },
        audio: false,
      });
      qrStreamRef.current = stream;
      if (!qrVideoRef.current) {
        setQrScanError("QR video element unavailable.");
        stopQrScanner();
        return;
      }
      qrVideoRef.current.srcObject = stream;
      await qrVideoRef.current.play();
      setQrScanning(true);

      const detector = new BarcodeDetector({ formats: ["qr_code"] });
      const scan = async () => {
        if (!qrVideoRef.current || !qrStreamRef.current) return;
        try {
          const results = await detector.detect(qrVideoRef.current);
          if (results.length) {
            const code = results[0].rawValue || "";
            if (activeIntakeTab === "guestSignOut") {
              updateSignOut("lookupValue", code);
            } else {
              updateCheckIn("lookupValue", code);
            }
            setQrScanError("");
            logAudit(
              "qr.scan.success",
              `QR code scanned for ${
                activeIntakeTab === "guestSignOut"
                  ? "guest sign-out"
                  : "returning guest"
              }`,
            );
            stopQrScanner();
            return;
          }
        } catch (err) {
          setQrScanError("Unable to read QR code. Try again.");
          stopQrScanner();
          return;
        }
        qrRafRef.current = requestAnimationFrame(scan);
      };

      qrRafRef.current = requestAnimationFrame(scan);
    } catch {
      setQrScanError("Camera access blocked. Allow camera permissions.");
      stopQrScanner();
    }
  }

  function stopQrScanner() {
    if (qrRafRef.current) {
      cancelAnimationFrame(qrRafRef.current);
      qrRafRef.current = null;
    }
    const stream = qrStreamRef.current;
    if (stream) {
      stream.getTracks().forEach((track) => track.stop());
    }
    qrStreamRef.current = null;
    if (qrVideoRef.current) {
      qrVideoRef.current.srcObject = null;
    }
    setQrScanning(false);
  }

  function validatePreReg() {
    const labels =
      REQUIRED_FIELD_LABELS[language] || REQUIRED_FIELD_LABELS.English;
    const missing = REQUIRED_PRE_REG_FIELDS.filter(
      (field) => !String(preReg[field] ?? "").trim(),
    );
    if (missing.length) {
      return `${t("missingPrefix")}: ${missing.map((field) => labels[field]).join(", ")}.`;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(preReg.visitorEmail)) {
      return t("invalidEmail");
    }
    if (preReg.reason === OTHER_REASON_VALUE && !preReg.reasonCustom.trim()) {
      return "Please specify a reason for the visit.";
    }
    if (!photoDataUrl) {
      return t("photoRequired");
    }
    return null;
  }

  function createPreRegistration() {
    const error = validatePreReg();
    if (error) {
      alert(error);
      return;
    }

    const visitorFullName = buildVisitorFullName(preReg);
    const visitReason =
      preReg.reason === OTHER_REASON_VALUE
        ? preReg.reasonCustom.trim()
        : preReg.reason;
    const host = employees.find((entry) => entry.id === preReg.hostId);
    if (!host) {
      alert("Selected host is not active.");
      return;
    }

    const visitStart = new Date(
      `${preReg.visitDate}T${preReg.visitTime}`,
    ).toISOString();
    const endAt = new Date(
      new Date(visitStart).getTime() +
        Number(preReg.expectedDurationMinutes || 60) * 60000,
    ).toISOString();

    const nowIso = new Date().toISOString();
    const record = {
      id: createVisitId(),
      qrCode: createToken(),
      pin: createPin(),
      timeToken: createToken(),
      verificationOtp: createPin(),
      createdAt: nowIso,
      host,
      visitorFullName,
      ...preReg,
      reason: visitReason,
      visitStart,
      endAt,
      checkInMethod: "Front Desk Sign-In",
      checkInLocation: "Reception-Desk",
      screening: {
        status: "clear",
        label: "Clear",
        detail: "No watchlist match.",
      },
      notification: { teams: false, email: false, sms: false, push: false },
      status: "checked_in",
      checkedInAt: nowIso,
      checkedOutAt: null,
      evacuatedAt: null,
      photoDataUrl,
      signature: "",
      mapPoint: "Main Lobby -> Elevator Bank B -> Security Turnstile",
    };

    // Reserve popup during click to avoid browser popup blocking.
    const reservedPrintWindow = window.open(
      "",
      "_blank",
      "width=900,height=900",
    );

    setRecords((prev) => [record, ...prev]);
    setActiveVisitId(record.id);
    setPreReg(getDefaultPreReg());
    clearPhoto();
    logAudit(
      "visit.preregistered",
      `Visit ${record.id} pre-registered for ${record.visitorFullName}`,
    );
    const printed = printBadgeForRecord(record, reservedPrintWindow);
    notifyHostForRecord(record).then((notified) => {
      const printMessage = printed
        ? "Badge print started."
        : "Badge print could not start (popup blocked).";
      const notifyMessage = notified
        ? "Host notification sent."
        : "Host notification queued locally.";
      showCompletionAlert(
        `${t("preRegDone")} ${t("visitId")}: ${record.id}\n${t("qrToken")}: ${record.qrCode}\n${t("pin")}: ${record.pin}\n${printMessage}\n${notifyMessage}`,
      );
    });
  }

  function findRecordForCheckIn() {
    const value = checkIn.lookupValue.trim();
    if (!value) return null;

    switch (checkIn.method) {
      case "QR":
        return records.find((entry) => entry.qrCode === value);
      case "PIN":
        return records.find((entry) => entry.pin === value);
      case "OTP":
        return records.find((entry) => entry.verificationOtp === value);
      case "Manual Lookup":
        return records.find(
          (entry) =>
            entry.visitorEmail.toLowerCase() === value.toLowerCase() ||
            entry.id === value,
        );
      case "Facial (Optional)":
        return records.find(
          (entry) =>
            entry.id === value ||
            entry.visitorFullName.toLowerCase() === value.toLowerCase(),
        );
      default:
        return null;
    }
  }

  function findRecordForSignOut() {
    const value = signOut.lookupValue.trim();
    if (!value) return null;

    switch (signOut.method) {
      case "QR":
        return records.find((entry) => entry.qrCode === value);
      case "PIN":
        return records.find((entry) => entry.pin === value);
      case "OTP":
        return records.find((entry) => entry.verificationOtp === value);
      case "Manual Lookup":
        return records.find(
          (entry) =>
            entry.visitorEmail.toLowerCase() === value.toLowerCase() ||
            entry.id === value,
        );
      case "Facial (Optional)":
        return records.find(
          (entry) =>
            entry.id === value ||
            entry.visitorFullName.toLowerCase() === value.toLowerCase(),
        );
      default:
        return null;
    }
  }

  useEffect(() => {
    const value = checkIn.lookupValue.trim();
    if (!value) {
      lastCompanyAutofillIdRef.current = null;
      if (!checkIn.company) {
        setCheckIn((prev) => ({ ...prev, company: DEFAULT_COMPANY }));
      }
      return;
    }

    const record = findRecordForCheckIn();
    if (!record) return;

    if (lastCompanyAutofillIdRef.current !== record.id) {
      lastCompanyAutofillIdRef.current = record.id;
      setCheckIn((prev) => ({
        ...prev,
        company: record.company || DEFAULT_COMPANY,
        verifyEmail: record.visitorEmail || "",
        verifyPhone: record.visitorPhone || "",
        verifyReason: record.reason || "",
        verifyHostId: record.hostId || record.host?.id || "",
        verifyDate: record.visitDate || "",
        verifyTime: record.visitTime || "",
        verifyDurationMinutes:
          record.expectedDurationMinutes != null
            ? String(record.expectedDurationMinutes)
            : "",
        verifyAccessLevel: record.accessLevel || "",
        verifyParking: record.parking || "",
        verifyAccommodations: record.accommodations || "",
      }));
    }
  }, [checkIn.lookupValue, checkIn.method, records]);

  function validateCheckInRecord(record) {
    if (!record) return "Visit not found for selected check-in method.";
    if (emergencyMode)
      return "Emergency mode is active. New check-ins are locked.";
    if (record.status === "checked_in") return "Visitor already checked in.";
    if (record.status === "checked_out") return "Visit already closed.";

    const now = new Date();
    const start = new Date(record.visitStart);
    const end = new Date(record.endAt);
    if (
      now < new Date(start.getTime() - 60 * 60000) ||
      now > new Date(end.getTime() + 60 * 60000)
    ) {
      return "Visit date/time window is not currently valid.";
    }

    if (
      checkIn.verifyEmail.trim().toLowerCase() !==
      record.visitorEmail.toLowerCase()
    ) {
      return "Email verification failed.";
    }

    if (
      checkIn.verifyPhone &&
      record.visitorPhone &&
      checkIn.verifyPhone.trim() !== record.visitorPhone.trim()
    ) {
      return "Phone verification failed.";
    }

    if (
      checkIn.verifyReason.trim().toLowerCase() !== record.reason.toLowerCase()
    ) {
      return "Reason verification failed.";
    }

    if (record.ndaRequired && !checkIn.signature.trim()) {
      return "Digital signature is required for NDA-enabled visits.";
    }

    if (!photoDataUrl) {
      return "Visitor photo capture is required.";
    }

    return null;
  }

  function performCheckIn() {
    const record = findRecordForCheckIn();
    const validationError = validateCheckInRecord(record);
    if (validationError) {
      alert(validationError);
      return;
    }

    const screening = checkWatchlist(record.visitorFullName);
    setScreeningResult(screening);

    if (screening.status === "review") {
      setRecords((prev) =>
        prev.map((entry) =>
          entry.id === record.id
            ? {
                ...entry,
                status: "security_hold",
                screening,
              }
            : entry,
        ),
      );
      logAudit("security.watchlist_hit", `Check-in blocked for ${record.id}`);
      alert("Watchlist match detected. Security and host notified.");
      return;
    }

    const checkedInAt = new Date().toISOString();
    const checkedInRecord = {
      ...record,
      status: "checked_in",
      checkedInAt,
      checkInMethod: checkIn.method,
      checkInLocation: checkIn.kioskId,
      screening,
      photoDataUrl,
      signature: checkIn.signature,
      company: checkIn.company || record.company,
    };

    // Reserve popup during the click event to avoid browser popup blocking.
    const reservedPrintWindow = window.open(
      "",
      "_blank",
      "width=900,height=900",
    );

    setRecords((prev) =>
      prev.map((entry) => (entry.id === record.id ? checkedInRecord : entry)),
    );
    setActiveVisitId(record.id);
    setCheckIn(EMPTY_CHECKIN);
    clearPhoto();
    logAudit(
      "visit.checked_in",
      `Visitor ${record.id} checked in via ${checkIn.method}`,
    );

    const printed = printBadgeForRecord(checkedInRecord, reservedPrintWindow);
    notifyHostForRecord(checkedInRecord).then((notified) => {
      const printMessage = printed
        ? "Badge print started."
        : "Badge print could not start (popup blocked).";
      const notifyMessage = notified
        ? "Host notification sent."
        : "Host notification queued locally.";
      showCompletionAlert(
        `Check-in complete.\n${printMessage}\n${notifyMessage}`,
      );
    });
  }

  function checkOutVisit(id, mode = "manual") {
    setRecords((prev) =>
      prev.map((entry) =>
        entry.id === id && entry.status === "checked_in"
          ? {
              ...entry,
              status: "checked_out",
              checkedOutAt: new Date().toISOString(),
            }
          : entry,
      ),
    );
    logAudit("visit.checked_out", `Visit ${id} checked out (${mode})`);
  }

  function validateSignOutRecord(record) {
    if (!record) return "Visit not found for selected sign-out method.";
    if (record.status === "checked_out") return "Visit already signed out.";
    if (record.status !== "checked_in") {
      return "Only active checked-in visits can be signed out.";
    }
    return null;
  }

  function performSignOut() {
    const record = findRecordForSignOut();
    const validationError = validateSignOutRecord(record);
    if (validationError) {
      alert(validationError);
      return;
    }

    setRecords((prev) =>
      prev.map((entry) =>
        entry.id === record.id
          ? {
              ...entry,
              status: "checked_out",
              checkedOutAt: new Date().toISOString(),
              checkOutMethod: signOut.method,
              checkOutLocation: signOut.kioskId,
            }
          : entry,
      ),
    );
    setActiveVisitId(record.id);
    setSignOut(EMPTY_SIGNOUT);
    logAudit(
      "visit.checked_out",
      `Visit ${record.id} checked out via ${signOut.method}`,
    );
    showCompletionAlert(t("signOutDone"));
  }

  function performManualSignOut(record) {
    const validationError = validateSignOutRecord(record);
    if (validationError) {
      alert(validationError);
      return;
    }

    setRecords((prev) =>
      prev.map((entry) =>
        entry.id === record.id
          ? {
              ...entry,
              status: "checked_out",
              checkedOutAt: new Date().toISOString(),
              checkOutMethod: "Manual Sign-Out",
              checkOutLocation: "Front Desk",
            }
          : entry,
      ),
    );
    setActiveVisitId(record.id);
    setSignOut(EMPTY_SIGNOUT);
    logAudit(
      "visit.checked_out",
      `Visit ${record.id} manually checked out from active guest list`,
    );
    showCompletionAlert(t("signOutDone"));
  }

  function autoExpireVisits() {
    const now = new Date();
    let changed = false;
    setRecords((prev) =>
      prev.map((entry) => {
        if (entry.status !== "checked_in") return entry;
        if (new Date(entry.endAt) < now) {
          changed = true;
          logAudit(
            "visit.auto_expired",
            `Visit ${entry.id} auto checked-out at expiry`,
          );
          return {
            ...entry,
            status: "checked_out",
            checkedOutAt: now.toISOString(),
          };
        }
        return entry;
      }),
    );
    if (!changed) {
      alert("No active visits reached expiration yet.");
    }
  }

  async function notifyHostForRecord(record) {
    const channels = Object.keys(notificationChannels).filter(
      (key) => notificationChannels[key],
    );
    if (!channels.length) {
      alert("Select at least one notification channel.");
      return false;
    }

    const payload = {
      event: "visitor_arrival",
      visitor: {
        name: record.visitorFullName,
        photo: Boolean(record.photoDataUrl),
        reason: record.reason,
        checkedInAt: record.checkedInAt,
        location: record.checkInLocation,
      },
      host: {
        name: record.host.name,
        email: record.host.email,
        phone: record.host.phone,
      },
      channels,
    };

    let delivered = true;
    try {
      const response = await fetch(NOTIFY_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!response.ok) throw new Error("notify failed");
    } catch {
      // Non-blocking; local demo fallback.
      delivered = false;
    }

    setRecords((prev) =>
      prev.map((entry) =>
        entry.id === record.id
          ? {
              ...entry,
              notification: {
                teams: entry.notification.teams || notificationChannels.teams,
                email: entry.notification.email || notificationChannels.email,
                sms: entry.notification.sms || notificationChannels.sms,
                push: entry.notification.push || notificationChannels.push,
              },
            }
          : entry,
      ),
    );
    logAudit(
      "host.notified",
      `Host notified for ${record.id} (${channels.join(",")})`,
    );
    return delivered;
  }

  async function notifyHost() {
    if (!activeRecord) {
      alert("Select a visit record first.");
      return;
    }
    const delivered = await notifyHostForRecord(activeRecord);
    alert(
      delivered
        ? "Host notification sent."
        : "Host notification queued locally.",
    );
  }

  function getBadgeQrValue(record) {
    return record.qrCode || record.id;
  }

  function createQrGeneratorTables() {
    const exp = new Array(512).fill(0);
    const log = new Array(256).fill(0);
    let value = 1;
    for (let i = 0; i < 255; i += 1) {
      exp[i] = value;
      log[value] = i;
      value <<= 1;
      if (value & 0x100) {
        value ^= 0x11d;
      }
    }
    for (let i = 255; i < 512; i += 1) {
      exp[i] = exp[i - 255];
    }
    return { exp, log };
  }

  const QR_TABLES = createQrGeneratorTables();

  function qrMultiply(a, b) {
    if (a === 0 || b === 0) return 0;
    return QR_TABLES.exp[QR_TABLES.log[a] + QR_TABLES.log[b]];
  }

  function createQrGeneratorPolynomial(degree) {
    let poly = [1];
    for (let i = 0; i < degree; i += 1) {
      const next = new Array(poly.length + 1).fill(0);
      for (let j = 0; j < poly.length; j += 1) {
        next[j] ^= qrMultiply(poly[j], QR_TABLES.exp[i]);
        next[j + 1] ^= poly[j];
      }
      poly = next;
    }
    return poly;
  }

  function createQrErrorCorrection(dataCodewords, degree) {
    const generator = createQrGeneratorPolynomial(degree);
    const message = [...dataCodewords, ...new Array(degree).fill(0)];
    for (let i = 0; i < dataCodewords.length; i += 1) {
      const factor = message[i];
      if (factor === 0) continue;
      for (let j = 1; j < generator.length; j += 1) {
        message[i + j] ^= qrMultiply(generator[j], factor);
      }
    }
    return message.slice(-degree);
  }

  function appendBits(target, value, length) {
    for (let i = length - 1; i >= 0; i -= 1) {
      target.push(((value >>> i) & 1) === 1);
    }
  }

  function buildQrCodewords(value) {
    const dataBytes = Array.from(new TextEncoder().encode(value));
    const dataCapacityBytes = 34;
    const dataCapacityBits = dataCapacityBytes * 8;
    if (dataBytes.length > 32) {
      throw new Error("Badge QR value is too long.");
    }

    const bits = [];
    appendBits(bits, 0b0100, 4);
    appendBits(bits, dataBytes.length, 8);
    dataBytes.forEach((byte) => appendBits(bits, byte, 8));
    appendBits(bits, 0, Math.min(4, dataCapacityBits - bits.length));

    while (bits.length % 8 !== 0) {
      bits.push(false);
    }

    const codewords = [];
    for (let i = 0; i < bits.length; i += 8) {
      let byte = 0;
      for (let j = 0; j < 8; j += 1) {
        byte = (byte << 1) | (bits[i + j] ? 1 : 0);
      }
      codewords.push(byte);
    }

    const padBytes = [0xec, 0x11];
    let padIndex = 0;
    while (codewords.length < dataCapacityBytes) {
      codewords.push(padBytes[padIndex % 2]);
      padIndex += 1;
    }

    const errorCorrection = createQrErrorCorrection(codewords, 10);
    return [...codewords, ...errorCorrection];
  }

  function getQrFormatBits(mask) {
    const data = (0b01 << 3) | mask;
    let rem = data;
    for (let i = 0; i < 10; i += 1) {
      rem = (rem << 1) ^ (((rem >>> 9) & 1) * 0x537);
    }
    return ((data << 10) | rem) ^ 0x5412;
  }

  function renderBadgeQrSvg(value) {
    const size = 25;
    const modules = Array.from({ length: size }, () => Array(size).fill(false));
    const isFunction = Array.from({ length: size }, () =>
      Array(size).fill(false),
    );

    function setFunctionModule(x, y, dark) {
      modules[y][x] = dark;
      isFunction[y][x] = true;
    }

    function drawFinderPattern(cx, cy) {
      for (let dy = -4; dy <= 4; dy += 1) {
        for (let dx = -4; dx <= 4; dx += 1) {
          const x = cx + dx;
          const y = cy + dy;
          if (x < 0 || x >= size || y < 0 || y >= size) continue;
          const dist = Math.max(Math.abs(dx), Math.abs(dy));
          setFunctionModule(x, y, dist !== 2 && dist !== 4);
        }
      }
    }

    function drawAlignmentPattern(cx, cy) {
      for (let dy = -2; dy <= 2; dy += 1) {
        for (let dx = -2; dx <= 2; dx += 1) {
          const dist = Math.max(Math.abs(dx), Math.abs(dy));
          setFunctionModule(cx + dx, cy + dy, dist !== 1);
        }
      }
    }

    function drawFormatBits(mask) {
      const bits = getQrFormatBits(mask);
      for (let i = 0; i <= 5; i += 1) {
        setFunctionModule(8, i, ((bits >>> i) & 1) !== 0);
      }
      setFunctionModule(8, 7, ((bits >>> 6) & 1) !== 0);
      setFunctionModule(8, 8, ((bits >>> 7) & 1) !== 0);
      setFunctionModule(7, 8, ((bits >>> 8) & 1) !== 0);
      for (let i = 9; i < 15; i += 1) {
        setFunctionModule(14 - i, 8, ((bits >>> i) & 1) !== 0);
      }
      for (let i = 0; i < 8; i += 1) {
        setFunctionModule(size - 1 - i, 8, ((bits >>> i) & 1) !== 0);
      }
      for (let i = 8; i < 15; i += 1) {
        setFunctionModule(8, size - 15 + i, ((bits >>> i) & 1) !== 0);
      }
      setFunctionModule(8, size - 8, true);
    }

    drawFinderPattern(3, 3);
    drawFinderPattern(size - 4, 3);
    drawFinderPattern(3, size - 4);
    drawAlignmentPattern(18, 18);

    for (let i = 8; i < size - 8; i += 1) {
      setFunctionModule(6, i, i % 2 === 0);
      setFunctionModule(i, 6, i % 2 === 0);
    }

    for (let i = 0; i < 8; i += 1) {
      if (i !== 6) {
        isFunction[8][i] = true;
        isFunction[i][8] = true;
      }
    }
    for (let i = 0; i < 7; i += 1) {
      isFunction[8][size - 1 - i] = true;
      isFunction[size - 1 - i][8] = true;
    }
    isFunction[8][8] = true;
    isFunction[8][7] = true;
    isFunction[7][8] = true;

    const allBits = buildQrCodewords(value).flatMap((byte) => {
      const out = [];
      appendBits(out, byte, 8);
      return out;
    });

    let bitIndex = 0;
    let upward = true;
    for (let right = size - 1; right >= 1; right -= 2) {
      if (right === 6) {
        right -= 1;
      }
      for (let vert = 0; vert < size; vert += 1) {
        const y = upward ? size - 1 - vert : vert;
        for (let j = 0; j < 2; j += 1) {
          const x = right - j;
          if (isFunction[y][x]) continue;
          modules[y][x] = allBits[bitIndex] || false;
          bitIndex += 1;
        }
      }
      upward = !upward;
    }

    for (let y = 0; y < size; y += 1) {
      for (let x = 0; x < size; x += 1) {
        if (!isFunction[y][x] && (x + y) % 2 === 0) {
          modules[y][x] = !modules[y][x];
        }
      }
    }

    drawFormatBits(0);

    const quietZone = 5;
    const rects = [];
    for (let y = 0; y < size; y += 1) {
      for (let x = 0; x < size; x += 1) {
        if (modules[y][x]) {
          rects.push(
            `<rect x="${x + quietZone}" y="${y + quietZone}" width="1" height="1" />`,
          );
        }
      }
    }

    return `
      <svg
        class="qr-svg"
        viewBox="0 0 ${size + quietZone * 2} ${size + quietZone * 2}"
        xmlns="http://www.w3.org/2000/svg"
        role="img"
        aria-label="Guest sign-out QR code"
      >
        <rect width="100%" height="100%" fill="#ffffff" />
        <g fill="#05131e">
          ${rects.join("")}
        </g>
      </svg>
    `;
  }

  function printBadgeForRecord(record, existingWindow = null) {
    const win =
      existingWindow || window.open("", "_blank", "width=900,height=900");
    if (!win) {
      return false;
    }

    const photo = record.photoDataUrl || "";
    const qrValue = getBadgeQrValue(record);
    const qrMarkup = renderBadgeQrSvg(qrValue);
    const html = `
      <!doctype html>
      <html>
        <head>
          <meta charset="utf-8" />
          <title>Visitor Badge</title>
          <style>
            @page { size: 4in 6in; margin: 0.2in; }
            * { box-sizing: border-box; }
            html, body { margin: 0; min-height: 100%; }
            body { font-family: Arial, sans-serif; padding: 24px; color: #05131e; background: #f0f6fb; }
            .page { min-height: calc(100vh - 48px); display: grid; place-items: center; }
            .badge { border: 2px solid #0f2f48; border-radius: 16px; padding: 14px; background: #ffffff; width: 3.6in; min-height: 5.55in; margin: 0 auto; display: grid; grid-template-rows: auto auto auto auto auto 1fr auto; gap: 8px; overflow: visible; }
            .v { margin: 0; font-size: 30px; font-weight: 800; letter-spacing: .08em; color: #0f2f48; text-align: center; }
            .name { margin: 0; font-size: 22px; font-weight: 700; line-height: 1.15; text-align: center; }
            .line { margin: 0; font-size: 12px; line-height: 1.3; }
            .photo-wrap { display: grid; place-items: center; min-height: 0; }
            img { width: 100%; max-width: 2.5in; aspect-ratio: 1 / 1; object-fit: cover; border-radius: 8px; border: 1px solid #c8d9e6; display: block; }
            .qr-wrap { display: grid; gap: 6px; justify-items: center; border: 0; padding: 16px 16px 12px; border-radius: 12px; overflow: visible; background: #ffffff; }
            .qr-svg { width: 1in; height: 1in; display: block; overflow: visible; }
            .qr-label { margin: 0; font-size: 9px; font-weight: 700; letter-spacing: .06em; color: #335b76; }
            .qr-token { margin: 0; font-size: 8px; line-height: 1.2; text-align: center; word-break: break-all; max-width: 100%; }
            @media print {
              html, body { width: 4in; height: 6in; min-height: 0; overflow: hidden; }
              body { background: #ffffff; padding: 0.2in; }
              .page { min-height: auto; display: block; }
              .badge { width: 3.6in; min-height: 5.55in; page-break-inside: avoid; break-inside: avoid; }
            }
          </style>
        </head>
        <body>
          <main class="page">
            <section class="badge">
              <p class="v">VISITOR</p>
              <p class="name">${record.visitorFullName}</p>
              <p class="line">Host: ${record.host.name}</p>
              <p class="line">Date: ${new Date(record.visitStart).toLocaleDateString()}</p>
              <p class="line">Expires: ${fmtTime(record.endAt)}</p>
              <p class="line">Visit ID: ${record.id}</p>
              ${
                photo
                  ? `<div class="photo-wrap"><img id="badge-photo" src="${photo}" alt="Visitor photo" /></div>`
                  : ""
              }
              <div class="qr-wrap">
                ${qrMarkup}
                <p class="qr-label">SCAN TO SIGN OUT</p>
                <p class="qr-token">${qrValue}</p>
              </div>
            </section>
          </main>
          <script>
            const badgePhoto = document.getElementById("badge-photo");

            async function waitForImage(img) {
              if (!img) {
                return;
              }

              try {
                if (typeof img.decode === "function") {
                  await img.decode();
                } else if (!img.complete) {
                  await new Promise((resolve) => {
                    img.addEventListener("load", resolve, { once: true });
                    img.addEventListener("error", resolve, { once: true });
                  });
                }
              } catch {
                // Allow printing even if decoding is unsupported or fails.
              }
            }

            async function prepareBadge() {
              await waitForImage(badgePhoto);
              window.__badgeReady = true;
            }

            if (document.readyState === "complete") {
              prepareBadge();
            } else {
              window.addEventListener("load", prepareBadge, { once: true });
            }

            window.addEventListener("afterprint", () => {
              window.close();
            }, { once: true });
          </script>
        </body>
      </html>
    `;

    win.document.open();
    win.document.write(html);
    win.document.close();
    let attempts = 0;
    const tryPrint = () => {
      attempts += 1;
      try {
        if (win.closed) {
          return;
        }
        if (win.__badgeReady) {
          win.focus();
          win.print();
          return;
        }
      } catch {
        return;
      }

      if (attempts < 60) {
        setTimeout(tryPrint, 100);
      }
    };
    setTimeout(tryPrint, 100);
    logAudit("badge.printed", `Badge printed for ${record.id}`);
    return true;
  }

  function showCompletionAlert(message) {
    setTimeout(() => {
      alert(message);
    }, 500);
  }

  function printBadge() {
    if (!activeRecord) {
      alert("Select a record before printing badge.");
      return;
    }
    const printed = printBadgeForRecord(activeRecord);
    alert(
      printed
        ? "Badge print started."
        : "Badge print could not start (popup blocked).",
    );
  }

  function toggleEmergency() {
    setEmergencyMode((prev) => {
      const next = !prev;
      logAudit(
        "emergency.mode",
        `Emergency mode ${next ? "enabled" : "disabled"}`,
      );
      return next;
    });
  }

  function markEvacuated(id) {
    setRecords((prev) =>
      prev.map((entry) =>
        entry.id === id
          ? {
              ...entry,
              evacuatedAt: entry.evacuatedAt ? null : new Date().toISOString(),
            }
          : entry,
      ),
    );
    logAudit("emergency.roll_call", `Roll-call updated for ${id}`);
  }

  function exportCsv() {
    if (!records.length) {
      alert("No records to export.");
      return;
    }

    const headers = [
      "Visit ID",
      "Visitor Name",
      "Email",
      "Phone",
      "Company",
      "Host",
      "Department",
      "Visit Start",
      "Expected End",
      "Status",
      "Check-In",
      "Check-Out",
      "Screening",
      "Location",
      "Check-Out Location",
      "Notification",
      "Evacuated At",
    ];

    const rows = records.map((entry) => [
      entry.id,
      entry.visitorFullName,
      entry.visitorEmail,
      entry.visitorPhone,
      entry.company,
      entry.host.name,
      entry.host.department,
      entry.visitStart,
      entry.endAt,
      entry.status,
      entry.checkedInAt || "",
      entry.checkedOutAt || "",
      entry.screening.label,
      entry.checkInLocation || "",
      entry.checkOutLocation || "",
      Object.entries(entry.notification)
        .filter(([, enabled]) => enabled)
        .map(([channel]) => channel)
        .join("+"),
      entry.evacuatedAt || "",
    ]);

    const csv = [headers, ...rows]
      .map((row) =>
        row
          .map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`)
          .join(","),
      )
      .join("\n");

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `vms-report-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function exportRollCall() {
    const active = records.filter((entry) => entry.status === "checked_in");
    if (!active.length) {
      alert("No active occupants.");
      return;
    }

    const headers = [
      "Visit ID",
      "Visitor",
      "Host",
      "Location",
      "Checked-In",
      "Evacuated",
    ];
    const rows = active.map((entry) => [
      entry.id,
      entry.visitorFullName,
      entry.host.name,
      entry.checkInLocation || "",
      entry.checkedInAt || "",
      entry.evacuatedAt || "",
    ]);

    const csv = [headers, ...rows]
      .map((row) =>
        row
          .map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`)
          .join(","),
      )
      .join("\n");

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `vms-rollcall-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  const signOutRecord = useMemo(
    () => findRecordForSignOut(),
    [signOut, records],
  );
  const signOutCandidates = useMemo(
    () =>
      records
        .filter((entry) => entry.status === "checked_in")
        .sort(
          (a, b) => new Date(b.checkedInAt || b.createdAt) - new Date(a.checkedInAt || a.createdAt),
        ),
    [records],
  );

  return (
    <main className="app shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">{t("eyebrow")}</p>
          <h1>{t("title")}</h1>
        </div>
        <div className="topbar-right">
          <p className="clock" aria-live="polite">
            {clock}
          </p>
          <button
            className="theme-toggle"
            type="button"
            onClick={() =>
              setTheme((prev) => (prev === "dark" ? "light" : "dark"))
            }
            aria-label={`Switch to ${theme === "dark" ? "light" : "dark"} mode`}
            title={`Switch to ${theme === "dark" ? "light" : "dark"} mode`}
          >
            {theme === "dark" ? (
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <circle cx="12" cy="12" r="4.2" fill="currentColor" />
                <g
                  stroke="currentColor"
                  strokeWidth="1.6"
                  strokeLinecap="round"
                >
                  <line x1="12" y1="2.5" x2="12" y2="5.2" />
                  <line x1="12" y1="18.8" x2="12" y2="21.5" />
                  <line x1="2.5" y1="12" x2="5.2" y2="12" />
                  <line x1="18.8" y1="12" x2="21.5" y2="12" />
                  <line x1="4.4" y1="4.4" x2="6.3" y2="6.3" />
                  <line x1="17.7" y1="17.7" x2="19.6" y2="19.6" />
                  <line x1="17.7" y1="6.3" x2="19.6" y2="4.4" />
                  <line x1="4.4" y1="19.6" x2="6.3" y2="17.7" />
                </g>
              </svg>
            ) : (
              <svg viewBox="0 0 24 24" aria-hidden="true">
                <path
                  d="M14.5 3a8.5 8.5 0 1 0 7.9 11.7A6.7 6.7 0 1 1 14.5 3Z"
                  fill="currentColor"
                />
              </svg>
            )}
          </button>
          <div className="prefs">
            <label>
              {t("role")}
              <select value={role} onChange={(e) => setRole(e.target.value)}>
                <option value="Admin">{t("admin")}</option>
                <option value="Front Desk">{t("frontDesk")}</option>
              </select>
            </label>
            <label>
              {t("language")}
              <select
                value={language}
                onChange={(e) => setLanguage(e.target.value)}
              >
                <option>English</option>
                <option>Spanish</option>
                <option>French</option>
              </select>
            </label>
          </div>
        </div>
      </header>

      {!isAdmin ? (
        <section className="panel intake-panel">
        <div className="intake-head">
          <h2>
            {activeIntakeTab === "guestSignIn"
              ? t("sectionPreReg")
              : activeIntakeTab === "returningGuest"
                ? t("sectionCheckIn")
                : t("sectionSignOut")}
          </h2>
          <div
            className="tab-switch"
            role="tablist"
            aria-label="Guest intake tabs"
          >
            <button
              className={`btn tab-btn ${activeIntakeTab === "guestSignIn" ? "primary" : ""}`}
              type="button"
              role="tab"
              aria-selected={activeIntakeTab === "guestSignIn"}
              onClick={() => setActiveIntakeTab("guestSignIn")}
            >
              {t("guestSignInTab")}
            </button>
            <button
              className={`btn tab-btn ${activeIntakeTab === "returningGuest" ? "primary" : ""}`}
              type="button"
              role="tab"
              aria-selected={activeIntakeTab === "returningGuest"}
              onClick={() => setActiveIntakeTab("returningGuest")}
            >
              {t("returningGuestTab")}
            </button>
            <button
              className={`btn tab-btn ${activeIntakeTab === "guestSignOut" ? "primary" : ""}`}
              type="button"
              role="tab"
              aria-selected={activeIntakeTab === "guestSignOut"}
              onClick={() => setActiveIntakeTab("guestSignOut")}
            >
              {t("guestSignOutTab")}
            </button>
          </div>
        </div>

        {activeIntakeTab === "guestSignIn" ? (
          <>
            <form className="form-grid" onSubmit={(e) => e.preventDefault()}>
              <label>
                {t("visitorFirstName")}
                <input
                  value={preReg.visitorFirstName}
                  onChange={(e) =>
                    updatePreReg("visitorFirstName", e.target.value)
                  }
                  placeholder="Jordan"
                />
              </label>
              <label>
                {t("visitorLastName")}
                <input
                  value={preReg.visitorLastName}
                  onChange={(e) =>
                    updatePreReg("visitorLastName", e.target.value)
                  }
                  placeholder="Lee"
                />
              </label>
              <label>
                {t("visitorEmail")}
                <input
                  type="email"
                  value={preReg.visitorEmail}
                  onChange={(e) => updatePreReg("visitorEmail", e.target.value)}
                  placeholder="jordan@company.com"
                />
              </label>
              <label>
                {t("visitorPhone")}
                <input
                  value={preReg.visitorPhone}
                  onChange={(e) =>
                    updatePreReg("visitorPhone", formatPhone(e.target.value))
                  }
                  placeholder="+1-555-0000"
                />
              </label>
              <label>
                {t("company")}
                <input
                  value={preReg.company}
                  onChange={(e) => updatePreReg("company", e.target.value)}
                  placeholder="Acme Inc"
                />
              </label>
              <label>
                {t("reasonForVisit")}
                <select
                  value={preReg.reason}
                  onChange={(e) => updatePreReg("reason", e.target.value)}
                >
                  <option value="">Select a reason</option>
                  {VISIT_REASONS.map((reason) => (
                    <option key={reason} value={reason}>
                      {reason}
                    </option>
                  ))}
                  <option value={OTHER_REASON_VALUE}>Other</option>
                </select>
              </label>
              {preReg.reason === OTHER_REASON_VALUE ? (
                <label className="span-2">
                  Explain Other
                  <input
                    value={preReg.reasonCustom}
                    onChange={(e) =>
                      updatePreReg("reasonCustom", e.target.value)
                    }
                    placeholder=""
                  />
                </label>
              ) : null}
              <label>
                {t("host")}
                <select
                  value={preReg.hostId}
                  onChange={(e) => updatePreReg("hostId", e.target.value)}
                >
                  <option value="">Select active employee</option>
                  {employees.map((emp) => (
                    <option key={emp.id} value={emp.id}>
                      {emp.name} - {emp.department}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t("date")}
                <input
                  type="date"
                  value={preReg.visitDate}
                  onChange={(e) => updatePreReg("visitDate", e.target.value)}
                />
              </label>
              <label>
                {t("time")}
                <input
                  type="time"
                  value={preReg.visitTime}
                  onChange={(e) => updatePreReg("visitTime", e.target.value)}
                />
              </label>
              <label>
                {t("expectedDuration")}
                <input
                  type="number"
                  min="15"
                  max="720"
                  value={preReg.expectedDurationMinutes}
                  onChange={(e) =>
                    updatePreReg("expectedDurationMinutes", e.target.value)
                  }
                />
              </label>
              <label>
                {t("accessLevel")}
                <select
                  value={preReg.accessLevel}
                  onChange={(e) => updatePreReg("accessLevel", e.target.value)}
                >
                  <option>Standard</option>
                  <option>Elevated</option>
                  <option>Restricted Escort</option>
                </select>
              </label>
              <label>
                {t("parking")}
                <input
                  value={preReg.parking}
                  onChange={(e) => updatePreReg("parking", e.target.value)}
                  placeholder="Lot B / Spot 22"
                />
              </label>
              <label>
                {t("ndaRequired")}
                <select
                  value={preReg.ndaRequired ? "yes" : "no"}
                  onChange={(e) =>
                    updatePreReg("ndaRequired", e.target.value === "yes")
                  }
                >
                  <option value="no">No</option>
                  <option value="yes">Yes</option>
                </select>
              </label>
              <label className="span-2">
                {t("specialAccommodations")}
                <textarea
                  rows="2"
                  value={preReg.accommodations}
                  onChange={(e) =>
                    updatePreReg("accommodations", e.target.value)
                  }
                  placeholder="Wheelchair support, interpreter, etc."
                />
              </label>
            </form>
            <article className="card">
              <div className="camera-head">
                <h3>{t("cameraTitle")}</h3>
                <p className="meta">{t("cameraHint")}</p>
              </div>
              <video
                ref={videoRef}
                autoPlay
                playsInline
                muted
                style={{ display: photoDataUrl ? "none" : "block" }}
              />
              <canvas ref={canvasRef} width="380" height="380" hidden />
              {photoDataUrl ? (
                <img
                  className="photo-preview"
                  src={photoDataUrl}
                  alt="visitor preview"
                />
              ) : null}
              <div className="row">
                <button
                  className="btn camera-preview-btn"
                  type="button"
                  onClick={startCamera}
                >
                  {cameraReady ? "Camera Ready" : "Start Camera"}
                </button>
                <button className="btn" type="button" onClick={stopCamera}>
                  Turn Off Camera
                </button>
                <button className="btn" type="button" onClick={capturePhoto}>
                  Capture
                </button>
                <button className="btn" type="button" onClick={clearPhoto}>
                  Retake
                </button>
              </div>
            </article>
            <div className="row">
              <button
                className="btn primary cta-lg"
                type="button"
                onClick={createPreRegistration}
              >
                {t("createPreReg")}
              </button>
            </div>
          </>
        ) : activeIntakeTab === "returningGuest" ? (
          <>
            <div className="form-grid">
              <label>
                {t("checkInMethod")}
                <select
                  value={checkIn.method}
                  onChange={(e) => updateCheckIn("method", e.target.value)}
                >
                  {CHECK_IN_METHODS.map((method) => (
                    <option key={method}>{method}</option>
                  ))}
                </select>
              </label>
            </div>

            {checkIn.method === "QR" ? (
              <article className="card qr-scan-card">
                <div className="camera-head">
                  <h3>QR Scan (Pre-Registered Guest)</h3>
                  <p className="meta">
                    Point the kiosk camera at the QR code on the visitor phone.
                  </p>
                </div>
                <video
                  ref={qrVideoRef}
                  autoPlay
                  playsInline
                  muted
                  className="qr-scan-video"
                  style={{ display: qrScanning ? "block" : "none" }}
                />
                {!qrScanning ? (
                  <div className="photo-placeholder">
                    Start scanning to read a QR code from a phone.
                  </div>
                ) : null}
                {qrScanError ? (
                  <p className="alert-line">{qrScanError}</p>
                ) : null}
                <div className="row">
                  <button
                    className="btn"
                    type="button"
                    onClick={startQrScanner}
                    disabled={qrScanning}
                  >
                    {qrScanning ? "Scanning..." : "Start QR Scan"}
                  </button>
                  <button
                    className="btn"
                    type="button"
                    onClick={stopQrScanner}
                    disabled={!qrScanning}
                  >
                    Stop Scan
                  </button>
                </div>
              </article>
            ) : null}

            <div className="form-grid">
              <label className="span-2 lookup-field">
                {t("lookupValue")}
                <input
                  type="search"
                  value={checkIn.lookupValue}
                  onChange={(e) => updateCheckIn("lookupValue", e.target.value)}
                />
              </label>
              <label className="span-2">
                {t("company")}
                <input
                  value={checkIn.company}
                  onChange={(e) => updateCheckIn("company", e.target.value)}
                />
              </label>
              <label>
                {t("confirmEmail")}
                <input
                  value={checkIn.verifyEmail}
                  onChange={(e) => updateCheckIn("verifyEmail", e.target.value)}
                />
              </label>
              <label>
                {t("confirmPhone")}
                <input
                  value={checkIn.verifyPhone}
                  onChange={(e) =>
                    updateCheckIn("verifyPhone", formatPhone(e.target.value))
                  }
                />
              </label>
              <label className="span-2">
                {t("confirmReason")}
                <input
                  value={checkIn.verifyReason}
                  onChange={(e) =>
                    updateCheckIn("verifyReason", e.target.value)
                  }
                />
              </label>
              <label>
                {t("confirmHost")}
                <select
                  value={checkIn.verifyHostId}
                  onChange={(e) =>
                    updateCheckIn("verifyHostId", e.target.value)
                  }
                >
                  <option value="">Select active employee</option>
                  {employees.map((emp) => (
                    <option key={emp.id} value={emp.id}>
                      {emp.name} - {emp.department}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t("confirmDate")}
                <input
                  type="date"
                  value={checkIn.verifyDate}
                  onChange={(e) => updateCheckIn("verifyDate", e.target.value)}
                />
              </label>
              <label>
                {t("confirmTime")}
                <input
                  type="time"
                  value={checkIn.verifyTime}
                  onChange={(e) => updateCheckIn("verifyTime", e.target.value)}
                />
              </label>
              <label>
                {t("confirmDuration")}
                <input
                  type="number"
                  min="15"
                  max="720"
                  value={checkIn.verifyDurationMinutes}
                  onChange={(e) =>
                    updateCheckIn("verifyDurationMinutes", e.target.value)
                  }
                />
              </label>
              <label>
                {t("confirmAccessLevel")}
                <select
                  value={checkIn.verifyAccessLevel}
                  onChange={(e) =>
                    updateCheckIn("verifyAccessLevel", e.target.value)
                  }
                >
                  <option value="">Select access level</option>
                  <option>Standard</option>
                  <option>Elevated</option>
                  <option>Restricted Escort</option>
                </select>
              </label>
              <label>
                {t("confirmParking")}
                <input
                  value={checkIn.verifyParking}
                  onChange={(e) =>
                    updateCheckIn("verifyParking", e.target.value)
                  }
                  placeholder="Lot B / Spot 22"
                />
              </label>
              <label className="span-2">
                {t("ndaSignature")}
                <input
                  value={checkIn.signature}
                  onChange={(e) => updateCheckIn("signature", e.target.value)}
                  placeholder="Type full legal name"
                />
              </label>
              <label className="span-2">
                {t("confirmAccommodations")}
                <textarea
                  rows="2"
                  value={checkIn.verifyAccommodations}
                  onChange={(e) =>
                    updateCheckIn("verifyAccommodations", e.target.value)
                  }
                  placeholder="Wheelchair support, interpreter, etc."
                />
              </label>
            </div>

            <article className="card">
              <div className="camera-head">
                <h3>{t("cameraTitle")}</h3>
                <p className="meta">{t("cameraHint")}</p>
              </div>
              <video
                ref={videoRef}
                autoPlay
                playsInline
                muted
                style={{ display: photoDataUrl ? "none" : "block" }}
              />
              <canvas ref={canvasRef} width="380" height="380" hidden />
              {photoDataUrl ? (
                <img
                  className="photo-preview"
                  src={photoDataUrl}
                  alt="visitor preview"
                />
              ) : null}
              <div className="row">
                <button
                  className="btn camera-preview-btn"
                  type="button"
                  onClick={startCamera}
                >
                  {cameraReady ? "Camera Ready" : "Start Camera"}
                </button>
                <button className="btn" type="button" onClick={stopCamera}>
                  Turn Off Camera
                </button>
                <button className="btn" type="button" onClick={capturePhoto}>
                  Capture
                </button>
                <button className="btn" type="button" onClick={clearPhoto}>
                  Retake
                </button>
              </div>
            </article>

            <div className="row">
              <button
                className="btn primary"
                type="button"
                onClick={performCheckIn}
                disabled={emergencyMode}
              >
                {t("completeCheckIn")}
              </button>
              <button className="btn" type="button" onClick={autoExpireVisits}>
                {t("runAutoExpiry")}
              </button>
            </div>
            {emergencyMode ? (
              <p className="alert-line">
                Emergency mode active: new check-ins locked.
              </p>
            ) : null}
          </>
        ) : (
          <>
            <div className="form-grid">
              <label>
                {t("signOutMethod")}
                <select
                  value={signOut.method}
                  onChange={(e) => updateSignOut("method", e.target.value)}
                >
                  {CHECK_IN_METHODS.map((method) => (
                    <option key={method}>{method}</option>
                  ))}
                </select>
              </label>
              <label>
                {t("kioskId")}
                <select
                  value={signOut.kioskId}
                  onChange={(e) => updateSignOut("kioskId", e.target.value)}
                >
                  {KIOSKS.map((kiosk) => (
                    <option key={kiosk}>{kiosk}</option>
                  ))}
                </select>
              </label>
            </div>

            {signOut.method === "QR" ? (
              <article className="card qr-scan-card">
                <div className="camera-head">
                  <h3>QR Scan (Guest Sign-Out)</h3>
                  <p className="meta">
                    Scan the QR code on the visitor badge or phone to sign out.
                  </p>
                </div>
                <video
                  ref={qrVideoRef}
                  autoPlay
                  playsInline
                  muted
                  className="qr-scan-video"
                  style={{ display: qrScanning ? "block" : "none" }}
                />
                {!qrScanning ? (
                  <div className="photo-placeholder">
                    Start scanning to locate the active visit record.
                  </div>
                ) : null}
                {qrScanError ? (
                  <p className="alert-line">{qrScanError}</p>
                ) : null}
                <div className="row">
                  <button
                    className="btn"
                    type="button"
                    onClick={startQrScanner}
                    disabled={qrScanning}
                  >
                    {qrScanning ? "Scanning..." : "Start QR Scan"}
                  </button>
                  <button
                    className="btn"
                    type="button"
                    onClick={stopQrScanner}
                    disabled={!qrScanning}
                  >
                    Stop Scan
                  </button>
                </div>
              </article>
            ) : null}

            <div className="form-grid">
              <label className="span-2 lookup-field">
                {t("signOutLookupValue")}
                <input
                  type="search"
                  value={signOut.lookupValue}
                  onChange={(e) => updateSignOut("lookupValue", e.target.value)}
                />
              </label>
            </div>

            <article className="card signout-summary-card">
              <div className="camera-head">
                <h3>{t("signOutSummary")}</h3>
                <span
                  className={`status ${signOutRecord?.status?.replaceAll("_", "-") || "waiting"}`}
                >
                  {signOutRecord
                    ? signOutRecord.status.replaceAll("_", " ")
                    : "waiting"}
                </span>
              </div>
              {signOutRecord ? (
                <div className="signout-summary-grid">
                  <p className="routing-host">{signOutRecord.visitorFullName}</p>
                  <p className="meta">{signOutRecord.id}</p>
                  <p className="meta">Host: {signOutRecord.host.name}</p>
                  <p className="meta">
                    {t("signOutStatusLabel")}:{" "}
                    {signOutRecord.status.replaceAll("_", " ")}
                  </p>
                  <p className="meta">
                    {t("signOutLocationLabel")}:{" "}
                    {signOutRecord.checkInLocation || "Unknown location"}
                  </p>
                  <p className="meta">
                    {t("signOutTimeLabel")}:{" "}
                    {fmtDateTime(signOutRecord.checkedInAt)}
                  </p>
                </div>
              ) : (
                <p className="meta">{t("signOutEmpty")}</p>
              )}
            </article>

            <article className="card signout-summary-card">
              <div className="camera-head">
                <h3>{t("signOutActiveList")}</h3>
                <span className="status checked-in">{signOutCandidates.length}</span>
              </div>
              {signOutCandidates.length ? (
                <div className="signout-candidates">
                  {signOutCandidates.map((entry) => (
                    <div
                      key={entry.id}
                      className={`signout-candidate ${
                        signOutRecord?.id === entry.id ? "selected" : ""
                      }`}
                    >
                      <div className="signout-candidate-copy">
                        <p className="routing-host">{entry.visitorFullName}</p>
                        <p className="meta">{entry.id}</p>
                        <p className="meta">Host: {entry.host.name}</p>
                        <p className="meta">
                          {t("signOutLocationLabel")}:{" "}
                          {entry.checkInLocation || "Unknown location"}
                        </p>
                        <p className="meta">
                          {t("signOutTimeLabel")}: {fmtDateTime(entry.checkedInAt)}
                        </p>
                      </div>
                      <button
                        className="btn mini"
                        type="button"
                        onClick={() => {
                          updateSignOut("lookupValue", getBadgeQrValue(entry));
                          setActiveVisitId(entry.id);
                        }}
                      >
                        {t("signOutSelect")}
                      </button>
                      <button
                        className="btn mini primary"
                        type="button"
                        onClick={() => performManualSignOut(entry)}
                      >
                        {t("completeSignOut")}
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="meta">{t("signOutActiveEmpty")}</p>
              )}
            </article>

            <div className="row">
              <button
                className="btn primary"
                type="button"
                onClick={performSignOut}
              >
                {t("completeSignOut")}
              </button>
            </div>
          </>
        )}
        </section>
      ) : null}

      {isAdmin ? (
        <section className="panel">
          <div className="list-head">
            <h2>{t("sectionAnalytics")}</h2>
            <div className="row tight">
              <button className="btn" type="button" onClick={exportCsv}>
                Export CSV
              </button>
            </div>
          </div>
          <div className="stats-grid">
            <article className="stat">
              <p>Daily Visitors</p>
              <strong>{metrics.dailyCount}</strong>
            </article>
            <article className="stat">
              <p>Weekly Trend</p>
              <strong>{metrics.weeklyCount}</strong>
            </article>
            <article className="stat">
              <p>Monthly Trend</p>
              <strong>{metrics.monthlyCount}</strong>
            </article>
            <article className="stat">
              <p>No-Show Rate</p>
              <strong>
                {records.length
                  ? `${Math.round((metrics.noShow / records.length) * 100)}%`
                  : "0%"}
              </strong>
            </article>
            <article className="stat">
              <p>Repeat Visitors</p>
              <strong>{metrics.repeat}</strong>
            </article>
            <article className="stat">
              <p>Average Visit (min)</p>
              <strong>{metrics.avgDuration}</strong>
            </article>
          </div>

          <section className="charts">
            <article className="card">
              <h3>Reason Distribution</h3>
              <div className="bars">
                {byReason.length ? (
                  byReason.map(([reason, count]) => (
                    <div className="bar-row" key={reason}>
                      <span>{reason}</span>
                      <div className="bar-track">
                        <div
                          className="bar-fill"
                          style={{
                            width: `${Math.max(8, (count / Math.max(...byReason.map((item) => item[1]), 1)) * 100)}%`,
                          }}
                        />
                      </div>
                      <strong>{count}</strong>
                    </div>
                  ))
                ) : (
                  <p className="meta">No data yet.</p>
                )}
              </div>
            </article>

            <article className="card">
              <h3>Time of Day Heat Bars</h3>
              <div className="hours-grid">
                {byHour.map((bucket) => (
                  <div
                    key={bucket.hour}
                    className="hour-col"
                    title={`${bucket.hour}:00 -> ${bucket.count}`}
                  >
                    <div
                      className="hour-fill"
                      style={{
                        height: `${Math.max(4, (bucket.count / Math.max(...byHour.map((item) => item.count), 1)) * 100)}%`,
                      }}
                    />
                    <span>{bucket.hour}</span>
                  </div>
                ))}
              </div>
            </article>
          </section>

          <section className="card records-card">
            <h3>{t("sectionRecords")}</h3>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Visit ID</th>
                    <th>Visitor</th>
                    <th>Host</th>
                    <th>Start</th>
                    <th>Status</th>
                    <th>Screening</th>
                    <th>Notifications</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {todaysRecords.length ? (
                    todaysRecords.map((entry) => (
                      <tr
                        key={entry.id}
                        className={activeVisitId === entry.id ? "selected" : ""}
                      >
                        <td>{entry.id}</td>
                        <td>{entry.visitorFullName}</td>
                        <td>{entry.host.name}</td>
                        <td>{fmtDateTime(entry.visitStart)}</td>
                        <td>
                          <span
                            className={`status ${entry.status.replaceAll("_", "-")}`}
                          >
                            {entry.status.replaceAll("_", " ")}
                          </span>
                        </td>
                        <td>{entry.screening.label}</td>
                        <td>
                          {Object.entries(entry.notification)
                            .filter(([, on]) => on)
                            .map(([name]) => name)
                            .join(" + ") || "Pending"}
                        </td>
                        <td>
                          <div className="row tight">
                            <button
                              className="btn mini"
                              type="button"
                              onClick={() => setActiveVisitId(entry.id)}
                            >
                              Select
                            </button>
                            <button
                              className="btn mini"
                              type="button"
                              onClick={() => checkOutVisit(entry.id, "manual")}
                              disabled={entry.status !== "checked_in"}
                            >
                              Check-Out
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td className="empty" colSpan="8">
                        No visit records for today.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        </section>
      ) : null}

      {isAdmin ? (
        <section className="layout">
          <section className="panel">
            <h2>{t("sectionHost")}</h2>
            <article className="card">
              <h3>Host Notification Channels</h3>
              <div className="row tight">
                {Object.keys(notificationChannels).map((channel) => (
                  <label key={channel} className="toggle">
                    <input
                      type="checkbox"
                      checked={notificationChannels[channel]}
                      onChange={(e) =>
                        setNotificationChannels((prev) => ({
                          ...prev,
                          [channel]: e.target.checked,
                        }))
                      }
                    />
                    {channel.toUpperCase()}
                  </label>
                ))}
              </div>
              <div className="row">
                <button className="btn" type="button" onClick={notifyHost}>
                  Send Arrival Notification
                </button>
                <button className="btn" type="button" onClick={printBadge}>
                  Print Physical Badge
                </button>
              </div>
            </article>

            <article className="card map-card">
              <h3>Meeting Map</h3>
              <div className="map-shell">
                <svg
                  viewBox="0 0 420 220"
                  className="map"
                  role="img"
                  aria-label="office meeting route map"
                >
                  <defs>
                    <linearGradient id="hallway" x1="0%" y1="0%" x2="100%" y2="100%">
                      <stop offset="0%" stopColor="#18364f" />
                      <stop offset="100%" stopColor="#0f2233" />
                    </linearGradient>
                  </defs>
                  <rect x="10" y="10" width="400" height="200" rx="20" fill="#0b1520" stroke="#315777" />
                  <rect x="28" y="28" width="110" height="58" rx="14" fill="#204560" />
                  <text x="46" y="61" fill="#d7efff" fontSize="13" fontWeight="700">Reception</text>
                  <text x="46" y="76" fill="#8eb7d8" fontSize="11">Guest arrival</text>
                  <rect x="156" y="42" width="110" height="50" rx="14" fill="#28516e" />
                  <text x="180" y="71" fill="#d7efff" fontSize="13" fontWeight="700">Security Desk</text>
                  <rect x="290" y="28" width="98" height="58" rx="14" fill="#204560" />
                  <text x="314" y="61" fill="#d7efff" fontSize="13" fontWeight="700">Elevators</text>
                  <text x="302" y="76" fill="#8eb7d8" fontSize="11">Floor access</text>
                  <rect x="56" y="126" width="148" height="54" rx="14" fill="url(#hallway)" />
                  <text x="82" y="157" fill="#d7efff" fontSize="13" fontWeight="700">Visitor Corridor</text>
                  <rect x="226" y="120" width="138" height="66" rx="14" fill="#16344b" />
                  <text x="249" y="151" fill="#d7efff" fontSize="13" fontWeight="700">Host Destination</text>
                  <text x="249" y="167" fill="#8eb7d8" fontSize="11">Meeting room / office</text>
                  <path d="M 139 57 C 147 57, 149 67, 156 67" fill="none" stroke="#7cc5ff" strokeWidth="6" strokeLinecap="round" />
                  <path d="M 266 67 C 274 67, 280 57, 289 57" fill="none" stroke="#7cc5ff" strokeWidth="6" strokeLinecap="round" />
                  <path d="M 340 88 C 340 110, 315 118, 295 120" fill="none" stroke="#7cc5ff" strokeWidth="6" strokeLinecap="round" />
                  <path d="M 226 153 C 214 153, 208 153, 204 153" fill="none" stroke="#7cc5ff" strokeWidth="6" strokeLinecap="round" />
                  <circle cx="84" cy="57" r="7" fill="#2dd4bf" />
                  <circle cx="210" cy="67" r="7" fill="#f59e0b" />
                  <circle cx="339" cy="57" r="7" fill="#38bdf8" />
                  <circle cx="295" cy="153" r="7" fill="#ef4444" />
                </svg>
                <div className="map-legend">
                  <span><i className="route-stop arrival" /> Arrival</span>
                  <span><i className="route-stop security" /> Security</span>
                  <span><i className="route-stop elevator" /> Elevator</span>
                  <span><i className="route-stop destination" /> Destination</span>
                </div>
              </div>
            </article>

            <article className="card">
              <h3>Recent Audit Events</h3>
              <ul className="feed compact">
                {audit.length ? (
                  audit
                    .slice(0, 8)
                    .map((entry) => (
                      <li
                        key={entry.id}
                      >{`${fmtTime(entry.at)} - ${entry.action} - ${entry.details}`}</li>
                    ))
                ) : (
                  <li>No audit events yet.</li>
                )}
              </ul>
            </article>

            <article className="card">
              <h3>Data Governance</h3>
              <label>
                Retention Period (Days)
                <input
                  type="number"
                  min="1"
                  max="3650"
                  value={retentionDays}
                  onChange={(e) =>
                    setRetentionDays(Number(e.target.value) || 90)
                  }
                />
              </label>
              <p className="meta">
                RBAC enabled. Audit logs include check-in/out, edits, watchlist,
                and emergency events.
              </p>
            </article>
          </section>

          <section className="panel">
            <h2>{t("sectionOccupancy")}</h2>
            <article className="card">
              <h3>Live Occupancy</h3>
              <p className="meta">In Building: {occupancy.length}</p>
              {occupancy.length ? (
                <div className="occupancy-list">
                  {Object.entries(
                    occupancy.reduce((acc, entry) => {
                      const floor =
                        entry.host.location.split("/")[1]?.trim() || "Unknown";
                      if (!acc[floor]) {
                        acc[floor] = [];
                      }
                      acc[floor].push(entry.visitorFullName);
                      return acc;
                    }, {}),
                  ).map(([floor, visitors]) => (
                    <div key={floor} className="occupancy-floor">
                      <p className="occupancy-floor-title">
                        {floor} ({visitors.length})
                      </p>
                      <p className="meta">{visitors.join(", ")}</p>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="meta">No active occupants.</p>
              )}
            </article>

            <article className="card">
              <h3>Emergency Mode</h3>
              <div className="row">
                <button
                  className={`btn ${emergencyMode ? "danger" : ""}`}
                  type="button"
                  onClick={toggleEmergency}
                >
                  {emergencyMode
                    ? "Disable Emergency Mode"
                    : "Enable Emergency Mode"}
                </button>
                <button className="btn" type="button" onClick={exportRollCall}>
                  Export Roll Call (CSV)
                </button>
              </div>
              <div className="emergency-grid">
                {occupancy.length ? (
                  occupancy.map((entry) => (
                    <article key={entry.id} className="emergency-card">
                      <div className="emergency-card-head">
                        <strong>{entry.visitorFullName}</strong>
                        <span className="status checked-in">On Site</span>
                      </div>
                      <p className="meta">{entry.id}</p>
                      <p className="meta">
                        {entry.checkInLocation || "Unknown location"}
                      </p>
                      <p className="meta">Host: {entry.host.name}</p>
                      <div className="row tight">
                        <button
                          className="btn mini"
                          type="button"
                          onClick={() => markEvacuated(entry.id)}
                        >
                          {entry.evacuatedAt ? "Undo" : "Mark Evacuated"}
                        </button>
                      </div>
                    </article>
                  ))
                ) : (
                  <div className="photo-placeholder emergency-empty">
                    No active occupants.
                  </div>
                )}
              </div>
            </article>

          </section>
        </section>
      ) : null}

    </main>
  );
}

export default App;
