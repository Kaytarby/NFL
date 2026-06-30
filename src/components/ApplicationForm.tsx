import React, { useState, useEffect, useRef, useMemo } from 'react';
import { Player, ApplicationDraft, fetchTeams, fetchZones, fetchPlayers, submitApplication, formatRussianDate } from '../lib/sheets';
import { uploadLogo } from '../lib/storage';
import { sendNotificationEmail } from '../lib/gmail';
import { Trophy, Upload, UserPlus, Check, X, AlertTriangle, Save, Loader2, Send, Shield, ListTodo, RefreshCw, Layers, Search, Users, Award, Flame, ChevronRight, Info, LayoutDashboard, Plus, ChevronDown, FileSpreadsheet, Printer, FileDown, FileText, Trash2 } from 'lucide-react';
import clsx from 'clsx';
import { getAuth } from 'firebase/auth';
import { saveSubmissionToFirestore, getSubmissionsFromFirestore, updateSubmissionData, deleteSubmissionFromFirestore, getTeamSubmission, FirestoreSubmission } from '../lib/firestore';
import * as XLSX from 'xlsx';

const DRAFT_KEY = 'nfl_application_draft';

interface ApplicationFormProps {
  onLogout: () => void;
  isGuest?: boolean;
  user?: any;
  defaultShowAdmin?: boolean;
  initialStage?: 'qualifier' | 'final';
}

export default function ApplicationForm({ onLogout, isGuest = false, user = null, defaultShowAdmin = false, initialStage }: ApplicationFormProps) {
  const [stage, setStage] = useState<'qualifier' | 'final' | null>(initialStage || null);
  const [adminStageTab, setAdminStageTab] = useState<'qualifier' | 'final'>('qualifier');
  const [loadingInitial, setLoadingInitial] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [draftSaving, setDraftSaving] = useState(false);
  
  const [showAdminPanel, setShowAdminPanel] = useState(defaultShowAdmin);
  const [zoneDates, setZoneDates] = useState<Record<string, string>>(() => {
    try {
      const saved = localStorage.getItem('NFL_ZONE_DATES');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [finalDate, setFinalDate] = useState<string>(() => {
    return localStorage.getItem('NFL_FINAL_DATE') || '01.06.2026';
  });
  const [deadlineDate, setDeadlineDate] = useState<string>(() => {
    return localStorage.getItem('NFL_DEADLINE_DATE') || '';
  });
  const [adminTab, setAdminTab] = useState<'drafts' | 'approved'>('drafts');
  const [openSubmissionId, setOpenSubmissionId] = useState<string | null>(null);
  const [expandedAdminZones, setExpandedAdminZones] = useState<Record<string, boolean>>({});
  const [submissions, setSubmissions] = useState<FirestoreSubmission[]>([]);
  const [loadingSubmissions, setLoadingSubmissions] = useState(false);
  const [syncingSubmissionId, setSyncingSubmissionId] = useState<string | null>(null);
  const [expandedSubmissions, setExpandedSubmissions] = useState<Record<string, boolean>>({});

  // Custom states and helpers for downloading rosters (Excel / print PDF)
  const [printSubmissions, setPrintSubmissions] = useState<FirestoreSubmission[]>([]);

  useEffect(() => {
    const handleAfterPrint = () => {
      setPrintSubmissions([]);
    };
    window.addEventListener('afterprint', handleAfterPrint);
    return () => window.removeEventListener('afterprint', handleAfterPrint);
  }, []);

  const handleExportAllToExcel = () => {
    const activeSubs = submissions.filter(s => (adminTab === 'approved' ? s.synced : !s.synced) && (s.stage || 'qualifier') === adminStageTab);
    if (activeSubs.length === 0) {
      alert("Нет заявок для экспорта");
      return;
    }

    const rows: any[] = [];
    activeSubs.forEach(sub => {
      sub.players.forEach(p => {
        const isLegString = p.isLegionnaire ? "Да" : "Нет";
        const isVerifiedString = p.isVerified ? "Заигран" : (p.fullName && dbPlayers.some(db => db.fullName.toLowerCase() === p.fullName?.toLowerCase()) ? "Был в базе" : "Новый");
        
        const rowData: any = {
          "Район": sub.zone || '-',
          "Команда": sub.teamName || '',
          "Капитан": sub.captainName || '',
          "Телефон": sub.captainPhone || '',
          "ФИО Игрока": p.fullName || '',
          "Дата рождения": p.birthDate || '',
          "Амплуа": p.position || '',
          "Игровой номер": p.number || '',
          "Легионер": isLegString,
          "Заигран": isVerifiedString
        };

        if (adminStageTab === 'final') {
           const mapStatus = { 'current': 'Текущий', 'new': 'Новый', 'other_zone': 'Др. Зона' };
           rowData["Статус (Финал)"] = p.transferStatus ? (mapStatus[p.transferStatus] || p.transferStatus) : 'Текущий';
        }

        rows.push(rowData);
      });
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Заявки");
    XLSX.writeFile(workbook, `NFL_${adminStageTab}_${adminTab === 'approved' ? 'Approved' : 'Drafts'}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleExportTeamToExcel = (sub: FirestoreSubmission) => {
    const rows = sub.players.map((p, idx) => {
      const isLegString = p.isLegionnaire ? "Да" : "Нет";
      const isVerifiedString = p.isVerified ? "Заигран" : (p.fullName && dbPlayers.some(db => db.fullName.toLowerCase() === p.fullName?.toLowerCase()) ? "В базе" : "Новый");
      
      const rowData: any = {
        "№": idx + 1,
        "ФИО Игрока": p.fullName || '',
        "Дата рождения": p.birthDate || '',
        "Амплуа": p.position || '',
        "Игровой номер": p.number || '',
        "Легионер": isLegString,
        "Заигран": isVerifiedString
      };

      if ((sub.stage || 'qualifier') === 'final') {
         const mapStatus = { 'current': 'Текущий', 'new': 'Новый', 'other_zone': 'Др. Зона' };
         rowData["Статус (Финал)"] = p.transferStatus ? (mapStatus[p.transferStatus] || p.transferStatus) : 'Текущий';
      }

      return rowData;
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Состав");
    XLSX.writeFile(workbook, `NFL_Team_${sub.teamName.replace(/\s+/g, '_')}_Excel.xlsx`);
  };

  const handlePrintAllToPdf = () => {
    const activeSubs = submissions.filter(s => (adminTab === 'approved' ? s.synced : !s.synced) && (s.stage || 'qualifier') === adminStageTab);
    if (activeSubs.length === 0) {
      alert("Нет заявок для печати");
      return;
    }
    setPrintSubmissions(activeSubs);
    setTimeout(() => {
      window.print();
    }, 500);
  };

  const handlePrintTeamToPdf = (sub: FirestoreSubmission) => {
    setPrintSubmissions([sub]);
    setTimeout(() => {
      window.print();
    }, 500);
  };

  const toggleSubmissionRoster = (submissionId: string) => {
    setExpandedSubmissions(prev => ({
      ...prev,
      [submissionId]: !prev[submissionId]
    }));
  };

  const [dbTeams, setDbTeams] = useState<string[]>([]);
  const [dbZones, setDbZones] = useState<string[]>([]);
  const [dbPlayers, setDbPlayers] = useState<Omit<Player, 'id'|'isConfirmed'|'status'>[]>([]);

  // Custom states added for validation popup & autocomplete & duplicate visual highlight
  const [errorModal, setErrorModal] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    steps: string[];
  }>({ isOpen: false, title: '', message: '', steps: [] });

  const [focusedPlayerInputId, setFocusedPlayerInputId] = useState<string | null>(null);
  const [showTeamSuggestions, setShowTeamSuggestions] = useState(false);
  const [showZoneSuggestions, setShowZoneSuggestions] = useState(false);
  const [selectedAdminTeamId, setSelectedAdminTeamId] = useState<string | null>(null);

  const [form, setForm] = useState<ApplicationDraft>({
    teamName: '',
    zone: '',
    captainName: '',
    captainPhone: '',
    logoUrl: null,
    players: []
  });

  const [dbLoadError, setDbLoadError] = useState<string | null>(null);
  const [logoFile, setLogoFile] = useState<File | null>(null);
  const [logoPreview, setLogoPreview] = useState<string | null>(null);

  useEffect(() => {
    const loadDb = async () => {
      let teams: string[] = [];
      let zones: string[] = [];
      let players: any[] = [];
      let loadErr: string | null = null;

      try {
        teams = await fetchTeams();
      } catch (err: any) {
        console.error("Fetch teams failed:", err);
        loadErr = err.message || String(err);
      }

      try {
        zones = await fetchZones();
      } catch (err: any) {
        console.error("Fetch zones failed:", err);
        if (!loadErr) loadErr = err.message || String(err);
      }

      try {
        players = await fetchPlayers();
      } catch (err: any) {
        console.error("Fetch players failed:", err);
        if (!loadErr) loadErr = err.message || String(err);
      }

      setDbTeams(teams);
      setDbZones(zones);
      setDbPlayers(players);

      // Prefill default tournament dates (01.06.2026) for loaded zones if not set
      setZoneDates(prev => {
        const next = { ...prev };
        let modified = false;
        zones.forEach(z => {
          if (!next[z]) {
            next[z] = '01.06.2026';
            modified = true;
          }
        });
        if (modified) {
          localStorage.setItem('NFL_ZONE_DATES', JSON.stringify(next));
        }
        return next;
      });

      if (loadErr) {
        setDbLoadError(loadErr);
      }

      // Load Draft
      try {
        const savedDraft = localStorage.getItem(DRAFT_KEY);
        if (savedDraft) {
          const parsed = JSON.parse(savedDraft);
          setForm(parsed);
          if (parsed.logoUrl) setLogoPreview(parsed.logoUrl);
        } else {
          // Pre-fill Captain Name from Google Account if possible
          const currentUser = getAuth().currentUser;
          if (currentUser && currentUser.displayName) {
             setForm(prev => ({ ...prev, captainName: currentUser.displayName || '' }));
          }
        }
      } catch (err) {
        console.error("Error loading draft", err);
      } finally {
        setLoadingInitial(false);
      }
    };
    loadDb();
  }, []);

  const loadSubmissions = async () => {
    try {
      setLoadingSubmissions(true);
      const data = await getSubmissionsFromFirestore();
      setSubmissions(data);
    } catch (err) {
      console.error("Failed to load submissions from Firestore", err);
    } finally {
      setLoadingSubmissions(false);
    }
  };

  useEffect(() => {
    loadSubmissions();
  }, []);

  const handleSyncToSheets = async (submission: FirestoreSubmission, sheetName: string) => {
    if (!submission.id) return;
    try {
      setSyncingSubmissionId(submission.id);
      
      const payload: ApplicationDraft = {
        teamName: submission.teamName,
        zone: submission.zone,
        captainName: submission.captainName,
        captainPhone: submission.captainPhone,
        logoUrl: submission.logoUrl,
        players: submission.players,
        version: getSubmissionVersion(submission)
      };

      // 1. Submit to Google Sheets!
      await submitApplication(payload, sheetName);
      
      // 2. Send automated Gmail notification!
      try {
        await sendNotificationEmail(submission.teamName, submission.captainName, submission.captainPhone);
      } catch (gmailErr) {
        console.warn("Gmail notification failed, sheet update succeeded:", gmailErr);
      }

      // 3. Mark as synced in Firestore!
      await updateSubmissionData(submission.id, { synced: true, status: 'approved' });

      alert(`Заявка команды "${submission.teamName}" успешно синхронизирована с Google Таблицами (Вкладка: ${sheetName})!`);
      
      // 4. Refresh submissions from Firestore
      await loadSubmissions();
    } catch (err: any) {
      console.error("Sync to sheets failed", err);
      alert(`Ошибка синхронизации: ${err.message}`);
    } finally {
      setSyncingSubmissionId(null);
    }
  };

  const handleRevokeApproval = async (submission: FirestoreSubmission) => {
    if (!submission.id) return;
    if (!window.confirm(`Вы уверены, что хотите отозвать статус "согласованной" у заявки команды "${submission.teamName}"?`)) return;
    try {
      setSyncingSubmissionId(submission.id);
      await updateSubmissionData(submission.id, { synced: false, status: 'pending' });
      await loadSubmissions();
      // NOTE: We don't delete rows from Google Sheets automatically to avoid complex row matching logic. 
      // The new submission sheet holds the history.
    } catch (err: any) {
      console.error("Revoke failed", err);
      alert("Ошибка при отзыве заявки: " + (err.message || err));
    } finally {
      setSyncingSubmissionId(null);
    }
  };

  const handleUpdateTeamName = async (id: string, newTeamName: string) => {
    try {
      const nameClean = newTeamName.trim().toLowerCase();
      const currentSub = submissions.find(s => s.id === id);
      const existing = submissions.find(s => s.id !== id && s.teamName.trim().toLowerCase() === nameClean && s.zone !== currentSub?.zone);
      
      if (existing) {
        alert(`Ошибка: команда с именем "${newTeamName}" уже существует в другом районе/зоне (${existing.zone}). Выберите уникальное имя.`);
        return;
      }
      
      await updateSubmissionData(id, { teamName: newTeamName });
      await loadSubmissions();
    } catch (err: any) {
      alert("Ошибка при переименовании: " + err.message);
    }
  };

  const handleDeleteSubmission = async (submission: FirestoreSubmission) => {
    if (!submission.id) return;
    if (window.confirm(`Вы уверены, что хотите удалить заявку команды "${submission.teamName}"? Это действие необратимо.`)) {
      try {
        await deleteSubmissionFromFirestore(submission.id);
        await loadSubmissions();
      } catch (err: any) {
        console.error("Failed to delete submission", err);
        alert(`Ошибка удаления: ${err.message}`);
      }
    }
  };

  const getSubmissionVersion = (sub: FirestoreSubmission) => {
    const teamSubs = submissions
      .filter(s => s.teamName.toLowerCase() === sub.teamName.toLowerCase())
      .sort((a, b) => {
        const timeA = a.createdAt ? new Date(a.createdAt).getTime() : 0;
        const timeB = b.createdAt ? new Date(b.createdAt).getTime() : 0;
        return timeA - timeB;
      });
    
    const index = teamSubs.findIndex(s => s.id === sub.id);
    return index !== -1 ? index + 1 : 1;
  };

  const isPlayerNewlyAdded = (sub: FirestoreSubmission, playerName: string) => {
    if (!sub.createdAt) return false;
    const subTime = new Date(sub.createdAt).getTime();

    const priorSubmissions = submissions.filter(s => 
      s.teamName.toLowerCase() === sub.teamName.toLowerCase() && 
      (s.stage || 'qualifier') === (sub.stage || 'qualifier') &&
      s.createdAt && new Date(s.createdAt).getTime() < subTime
    );

    const nameClean = playerName.trim().toLowerCase();

    if (priorSubmissions.length > 0) {
      for (const priorSub of priorSubmissions) {
        if (priorSub.players.some(p => p.fullName.trim().toLowerCase() === nameClean)) {
          return false;
        }
      }
      return true;
    }

    const existingTeamPlayers = dbPlayers.filter(p => p.teamName.trim().toLowerCase() === sub.teamName.trim().toLowerCase());
    
    if (existingTeamPlayers.some(p => p.fullName.trim().toLowerCase() === nameClean)) {
      return false;
    }

    return true;
  };

  const parseDDMMYYYYToTime = (dateStr: string) => {
    if (!dateStr || dateStr.length !== 10) return 0;
    const [d, m, y] = dateStr.split('.');
    return new Date(`${y}-${m}-${d}T23:59:59.999Z`).getTime(); 
  };

  const getPlayerApprovedTeam = (fullName: string | undefined, currentTeamName: string): string | null => {
    if (!fullName || !fullName.trim()) return null;
    const nameClean = fullName.trim().toLowerCase();
    
    const matchedSub = submissions.find(sub => {
      if (sub.teamName.toLowerCase() === currentTeamName.toLowerCase()) return false;
      if (!sub.synced && sub.status !== 'approved') return false;
      return (sub.players || []).some(p => p.fullName && p.fullName.trim().toLowerCase() === nameClean);
    });
    
    return matchedSub ? matchedSub.teamName : null;
  };

  const saveDraft = () => {
    setDraftSaving(true);
    localStorage.setItem(DRAFT_KEY, JSON.stringify(form));
    setTimeout(() => setDraftSaving(false), 500);
  };

  useEffect(() => {
    const delayDebounceFn = setTimeout(() => {
      if (!loadingInitial) saveDraft();
    }, 1000);
    return () => clearTimeout(delayDebounceFn);
  }, [form]);

  // Derive suggested teams including all submissions to help them find mismatched names
  const suggestedTeams = useMemo(() => {
    const fromDb = [...dbTeams];
    
    // Always include teams from all submissions
    const fromSubs = submissions
      .filter(s => s.teamName)
      // If zone is selected, only show teams from that zone, otherwise show all
      .filter(s => !form.zone || s.zone === form.zone)
      .map(s => s.teamName.trim());
      
    fromSubs.forEach(t => {
      if (!fromDb.find(d => d.toLowerCase() === t.toLowerCase())) {
        fromDb.push(t);
      }
    });

    return fromDb;
  }, [dbTeams, submissions, form.zone]);

  const handleTeamNameChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newTeamName = e.target.value;
    setForm(prev => ({ ...prev, teamName: newTeamName }));
  };

  const handleTeamSelect = async (name: string) => {
    // Try to fetch previous submission from local submissions state
    let prevSub = null;
    const nameClean = name.trim().toLowerCase();
    const zoneFilter = form.zone ? form.zone.trim().toLowerCase() : null;

    const matchedSubs = submissions.filter(s => {
       if (s.teamName.trim().toLowerCase() !== nameClean) return false;
       if (zoneFilter && s.zone && s.zone.trim().toLowerCase() !== zoneFilter) return false;
       return true;
    }).sort((a, b) => {
       const timeA = a.createdAt ? new Date(a.createdAt).getTime() : 0;
       const timeB = b.createdAt ? new Date(b.createdAt).getTime() : 0;
       return timeB - timeA; // desc
    });
    
    if (stage === 'final') {
      prevSub = matchedSubs.find(s => (s.stage || 'qualifier') === 'final');
      if (!prevSub) {
        prevSub = matchedSubs.find(s => (s.stage || 'qualifier') === 'qualifier');
      }
    } else {
      prevSub = matchedSubs.find(s => (s.stage || 'qualifier') === 'qualifier');
    }
    
    if (prevSub) {
      setForm(prev => ({
        ...prev,
        teamName: name,
        zone: prevSub?.zone || prev.zone,
        captainName: prevSub?.captainName || '',
        captainPhone: prevSub?.captainPhone || '',
        logoUrl: prevSub?.logoUrl || '',
        players: (prevSub?.players || []).filter(p => p.status !== 'deleted').map(p => ({
          ...p,
          id: Math.random().toString(36).substring(7),
          isConfirmed: true,
          status: 'previous' as const
        }))
      }));
      if (prevSub.logoUrl) setLogoPreview(prevSub.logoUrl);
      return;
    }

    setForm(prev => {
      // Find players for this team from Sheets baseline
      const existingTeamPlayers = dbPlayers.filter(p => p.teamName.trim().toLowerCase() === name.trim().toLowerCase());
      
      const newPlayersStr = existingTeamPlayers.map(p => ({
        ...p,
        id: Math.random().toString(36).substring(7),
        isConfirmed: true,
        status: 'previous' as const
      }));

      return {
        ...prev,
        teamName: name,
        players: newPlayersStr
      };
    });
  };

  const handleAddPlayer = () => {
    const newId = 'player-' + Math.random().toString(36).substring(7);
    setForm(prev => ({
      ...prev,
      players: [
        ...prev.players,
        {
          id: newId,
          teamName: prev.teamName,
          fullName: '',
          birthDate: '',
          position: '',
          number: '',
          isConfirmed: true,
          status: 'new'
        }
      ]
    }));

    setTimeout(() => {
      const el = document.getElementById(newId);
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    }, 150);
  };

  const removePlayer = (id: string, status: string) => {
    setForm(prev => ({
      ...prev,
      players: prev.players.map(p => {
        if (p.id === id) {
          if (p.status === 'previous') return { ...p, status: 'deleted', isConfirmed: false };
          return p;
        }
        return p;
      }).filter(p => !(p.id === id && p.status === 'new')) // physically remove if it's completely new
    }));
  };

  const confirmPlayer = (id: string) => {
    setForm(prev => ({
      ...prev,
      players: prev.players.map(p => p.id === id ? { ...p, isConfirmed: true, status: 'previous' } : p)
    }));
  };
  
  const restorePlayer = (id: string) => {
    setForm(prev => ({
      ...prev,
      players: prev.players.map(p => p.id === id ? { ...p, status: 'previous' } : p)
    }));
  };

  const updatePlayer = (id: string, field: keyof Player, value: any) => {
    setForm(prev => ({
      ...prev,
      players: prev.players.map(p => {
        if (p.id === id) {
          const newP = { ...p, [field]: value };
          // If they typed a name that matches DB exactly, autofill other fields
          if (field === 'fullName') {
             const exactMatch = dbPlayers.find(db => db.fullName.toLowerCase() === value.toLowerCase());
             if (exactMatch) {
               newP.birthDate = exactMatch.birthDate;
               newP.position = exactMatch.position;
               newP.number = exactMatch.number;
               newP.isLegionnaire = exactMatch.isLegionnaire;
               newP.isVerified = exactMatch.isVerified;
             }
          }
          return newP;
        }
        return p;
      })
    }));
  };

  const handleLogoChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      setLogoFile(file);
      setLogoPreview(URL.createObjectURL(file));
      // Reset the saved url since it's a new file
      setForm(prev => ({ ...prev, logoUrl: null }));
    }
  };

  const calculateAge = (birthDateStr: string, tournamentDateStr: string = '01.06.2026') => {
    const [bDay, bMo, bYr] = birthDateStr.split('.').map(Number);
    const [tDay, tMo, tYr] = tournamentDateStr.split('.').map(Number);
    if (!bDay || !bMo || !bYr || !tDay || !tMo || !tYr) return null;
    let age = tYr - bYr;
    if (tMo < bMo || (tMo === bMo && tDay < bDay)) {
      age--;
    }
    return age;
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    // 0. Check unique team name across other zones
    const teamNameClean = form.teamName.trim().toLowerCase();
    const existingOtherZone = submissions.find(
      s => s.teamName.trim().toLowerCase() === teamNameClean &&
           s.zone && s.zone !== form.zone
    );
    if (existingOtherZone) {
      setErrorModal({
        isOpen: true,
        title: "Название команды занято",
        message: `Команда с названием "${form.teamName}" уже зарегистрирована в районе "${existingOtherZone.zone}".`,
        steps: ["Пожалуйста, выберите уникальное название для вашей команды (например, добавьте название села или района)."]
      });
      return;
    }

    // 1. Check basic group fields
    if (!form.teamName.trim() || !form.captainName.trim() || !form.captainPhone.trim() || !form.zone.trim()) {
      setErrorModal({
        isOpen: true,
        title: "Неполные данные анкеты",
        message: "Для подачи заявки необходимо полностью указать сведения о команде и капитане.",
        steps: [
          "Введите название вашей футбольной команды",
          "Выберите или укажите соответствующий район (зону)",
          "Укажите полное ФИО капитана (контактного лица)",
          "Введите корректный номер телефона капитана для обратной связи"
        ]
      });
      return;
    }

    const activePlayers = form.players.filter(p => p.status !== 'deleted' && p.fullName.trim() !== '');

    // 2. Validate Legionnaires
    const legionnaires = activePlayers.filter(p => p.isLegionnaire);
    if (legionnaires.length > 1) {
       setErrorModal({
         isOpen: true,
         title: "Превышен лимит легионеров",
         message: `В заявке указано ${legionnaires.length} легионера(ов). Разрешен максимум 1 легионер на команду.`,
         steps: [
           "Уберите отметку Легионер у лишних игроков"
         ]
       });
       return;
    }

    // 3. Validate Age
    let under16 = [];
    let under18 = [];
    const tDate = stage === 'final' ? finalDate : (zoneDates[form.zone] || '01.06.2026');
    for (const p of activePlayers) {
      if (p.birthDate.length === 10) {
        const age = calculateAge(p.birthDate, tDate);
        if (age !== null && age < 16) {
          under16.push(p.fullName);
        } else if (age !== null && age < 18) {
          under18.push(p.fullName);
        }
      }
    }

    if (under16.length > 0) {
      setErrorModal({
         isOpen: true,
         title: "Возрастное ограничение",
         message: "Недопустимый возраст. Игрокам не должно быть менее 16 лет на дату проведения турнира.",
         steps: [
           `Уберите следующих игроков из заявки: ${under16.join(', ')}`
         ]
       });
       return;
    }

    // 4. Validate player count (min 10, max 15)
    const activeCount = activePlayers.length;

    if (activeCount < 10 || activeCount > 15) {
      setErrorModal({
        isOpen: true,
        title: "Неверное количество игроков в заявке",
        message: `По регламенту турнира НОГАЙСКОЙ ФУТБОЛЬНОЙ ЛИГИ в официальную заявку допускается заявлять строго от 10 до 15 футболистов. Ваш текущий состав содержит: ${activeCount} игроков.`,
        steps: [
          activeCount < 10 
            ? `Добавьте еще как минимум ${10 - activeCount} игроков, нажав на кнопку "Добавить нового игрока" и заполнив их ФИО и даты рождения.`
            : `Уберите лишних ${activeCount - 15} игроков, нажав "✖" на лишних карточках, чтобы уложиться в лимит 15 человек.`,
          "Вы можете удалять ненужных игроков или добавлять новых до тех пор, пока количество вбитых игроков не станет от 10 до 15.",
          "Игроки с пустыми именами не учитываются в общем подсчёте."
        ]
      });
      return;
    }

    // 5. Final Stage Transfer Rules
    if (stage === 'final') {
      const missingStatus = activePlayers.filter(p => !p.transferStatus);
      if (missingStatus.length > 0) {
        setErrorModal({
          isOpen: true,
          title: "Заполните статусы изменений",
          message: "При заявке на Финал необходимо обязательно указать статус (Текущий, Новый, Из другой зоны) у каждого футболиста состава.",
          steps: [
            "Проверьте, что во всех анкетах игроков заполнено поле 'Статус игрока относительно отборочного этапа'"
          ]
        });
        return;
      }
      
      const newPlayers = activePlayers.filter(p => p.transferStatus === 'new' || p.transferStatus === 'other_zone');
      const otherZone = activePlayers.filter(p => p.transferStatus === 'other_zone');

      if (newPlayers.length > 5 || otherZone.length > 3) {
        const errorMsg = [];
        if (newPlayers.length > 5) errorMsg.push(`Новых изменений всего заявлено ${newPlayers.length} (максимум разрешено 5).`);
        if (otherZone.length > 3) errorMsg.push(`Заиграно из другой зоны ${otherZone.length} (максимум разрешено 3).`);

        setErrorModal({
          isOpen: true,
          title: "Превышен лимит замен",
          message: "В Финале заявки действуют жесткие ограничения на изменения относительно основного отборочного состава.",
          steps: [
            ...errorMsg,
            "Измените статусы или уберите лишних 'новых' и 'других зон'."
          ]
        });
        return;
      }
    }

    const finalConfirmMsg = `Подтверждаете отправку заявки для команды "${form.teamName}" (заявлено ${activeCount} игроков) на сезон 2026?`;
    if (!window.confirm(finalConfirmMsg)) return;

    try {
      setSubmitting(true);
      
      // Prepare final submitted object with logo as null (logo upload removed by request)
      const draftVersion = submissions.filter(s => s.teamName.toLowerCase() === form.teamName.toLowerCase()).length + 1;
      
      const finalForm = {
        ...form,
        logoUrl: null,
        players: activePlayers,
        version: draftVersion,
        stage: stage || 'qualifier'
      };

      if (isGuest) {
        await saveSubmissionToFirestore(finalForm);
        alert("Заявка успешно принята и отправлена на подтверждение в оргкомитет НФЛ!");
      } else {
        await submitApplication(finalForm, 'Черновики');
        try {
          await sendNotificationEmail(form.teamName, form.captainName, form.captainPhone);
        } catch (gmailErr) {
          console.warn("Gmail notification failed, sheet update succeeded:", gmailErr);
        }
        alert("Заявка успешно отправлена напрямую в Google Таблицу (в лист Черновики)!");
      }

      localStorage.removeItem(DRAFT_KEY);
      
      // Reset form
      setForm({ teamName: '', zone: '', captainName: '', captainPhone: '', logoUrl: null, players: [] });
      setLogoFile(null);
      setLogoPreview(null);
      
    } catch (err: any) {
      console.error(err);
      setErrorModal({
        isOpen: true,
        title: "Ошибка при отправке заявки",
        message: `При отправке данных на сервер Файрбейз/Google Sheets произошла ошибка: ${err.message}`,
        steps: [
          "Проверьте стабильность интернет-соединения.",
          "Убедитесь, что заполнены все поля анкеты.",
          "Если проблема повторяется, обратитесь напрямую к организаторам лиги в оргкомитет."
        ]
      });
    } finally {
      setSubmitting(false);
    }
  };

  if (loadingInitial) {
    return (
      <div className="min-h-screen bg-[#020516] star-bg text-white flex flex-col items-center justify-center p-4">
        <Loader2 className="w-12 h-12 text-[#c5a85c] animate-spin mb-4" />
        <span className="text-lg tracking-wider font-display text-slate-300">Загрузка данных турнира...</span>
      </div>
    );
  }

  // Find autocomplete suggestions for players based on currently typed names
  const getPlayerSuggestions = (typed: string) => {
     if (!typed || typed.length < 2) return [];
     return dbPlayers
        .filter(p => p.fullName.toLowerCase().includes(typed.toLowerCase()))
        .slice(0, 5); // top 5
  };

  return (
    <>
      <div className={clsx("min-h-screen bg-[#020516] star-bg text-slate-100 font-sans pb-24 relative overflow-hidden", printSubmissions.length > 0 ? "hidden" : "block")}>
      
      {/* Champions League night ambient glow */}
      <div className="absolute top-0 right-0 w-[500px] h-[500px] bg-gradient-to-br from-blue-600/10 to-transparent rounded-full filter blur-[120px] pointer-events-none" />
      <div className="absolute bottom-1/4 left-0 w-[400px] h-[400px] bg-gradient-to-tr from-cyan-500/10 to-transparent rounded-full filter blur-[120px] pointer-events-none" />

      <header className="bg-slate-950/80 backdrop-blur-md sticky top-0 z-50 border-b border-white/5 shadow-2xl">
        <div className="max-w-4xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
             <div className="bg-slate-900 border border-slate-755 p-2 rounded-xl text-[#c5a85c] shadow-[0_0_15px_rgba(197,168,92,0.1)]">
                <Trophy className="w-5 h-5" />
             </div>
             <div>
               <h1 className="font-extrabold text-base md:text-lg tracking-tight leading-none text-white font-display">
                 Заявка на отбор
               </h1>
               <p className="text-xs font-semibold tracking-widest text-[#c5a85c] uppercase mt-1">
                 Ногайская Футбольная Лига
               </p>
             </div>
          </div>
          <div className="flex items-center gap-4">
             {draftSaving && (
               <span className="text-xs text-yellow-400 font-medium flex items-center gap-1.5 px-2 py-1 bg-yellow-400/10 rounded-full border border-yellow-400/20">
                 <Save className="w-3.5 h-3.5 text-[#c5a85c] animate-pulse" /> Черновик сохранен
               </span>
             )}
             {user && (
               <button
                 type="button"
                 onClick={() => setShowAdminPanel(!showAdminPanel)}
                 className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-gradient-to-r from-blue-700 to-blue-800 hover:from-blue-600 hover:to-blue-700 text-white font-extrabold rounded-lg transition-all border border-blue-500/20 cursor-pointer shadow-lg shadow-blue-500/10"
               >
                 <Shield className="w-3.5 h-3.5" />
                 {showAdminPanel ? "Форма заявки" : "Панель заявок"}
                 {submissions.filter(s => !s.synced).length > 0 && (
                   <span className="bg-red-500 text-white text-[9px] w-4.5 h-4.5 rounded-full flex items-center justify-center font-bold animate-pulse">
                     {submissions.filter(s => !s.synced).length}
                   </span>
                 )}
               </button>
             )}
             <button 
               onClick={onLogout} 
               className="text-xs md:text-sm font-semibold text-slate-400 hover:text-[#c5a85c] hover:underline underline-offset-4 transition-all"
             >
               Выйти
             </button>
          </div>
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-4 py-8 relative z-10">
        
        {dbLoadError && (
          <div className="mb-6 p-4 rounded-xl bg-red-950/40 border border-red-500/20 text-red-200 text-xs shadow-lg leading-relaxed relative overflow-hidden">
            <div className="absolute top-0 left-0 bottom-0 w-[4px] bg-red-500" />
            <div className="flex items-start gap-3 pl-1">
              <AlertTriangle className="w-5 h-5 text-red-400 flex-shrink-0 mt-0.5 animate-pulse" />
              <div className="space-y-1">
                <p className="font-extrabold text-sm text-red-300">⚠️ База данных Google Таблиц не загружена</p>
                <p className="text-slate-300 text-[11px]">
                  Не удалось загрузить списки команд и футболистов, из-за чего автозаполнение при вводе названия команды или ФИО игрока временно недоступно. Пожалуйста, выполните следующие настройки:
                </p>
                <ul className="list-disc pl-4 space-y-1.5 text-slate-400 text-[11px] mt-1.5">
                  <li>
                    <strong className="text-slate-300">Доступ к таблице:</strong> Откройте Google Таблицу на Google Диске, выберите <strong className="text-white">Поделиться / Share</strong> в верхнем правом углу, и в разделе общего доступа установите режим <strong className="text-[#c5a85c]">"Все, у кого есть ссылка" (Anyone with link can view)</strong> со статусом <strong className="text-emerald-400">"Читатель" (Viewer)</strong>. Без этого настройки приватности Google заблокируют чтение сайта.
                  </li>
                  <li>
                    <strong className="text-slate-300">Корректность Spreadsheet ID:</strong> Проверьте, добавлен ли правильный ID вашей Google-таблицы в переменную окружения <strong className="text-white">VITE_SPREADSHEET_ID</strong> в настройках проекта.
                  </li>
                  <li>
                    <strong className="text-slate-300">Названия вкладок (листов):</strong> Убедитесь, что внутри вашей таблицы вкладки на латинице названы ровно <strong className="text-white">'Teams'</strong>, <strong className="text-white">'Zones'</strong> и <strong className="text-white">'Players'</strong> (с заглавной буквы). Если они на кириллице (например, "Команды"), API не сможет их сопоставить.
                  </li>
                </ul>
                <div className="mt-3 bg-red-950/60 p-2.5 rounded border border-red-500/10 font-mono text-[10px] text-red-350 select-all flex flex-wrap gap-1">
                  <span className="font-bold text-red-400">Техническая ошибка:</span> <span>{dbLoadError}</span>
                </div>
              </div>
            </div>
          </div>
        )}

        {showAdminPanel && user ? (
          <div className="space-y-6">
            
            {/* Header banner */}
            <div className="p-6 md:p-8 rounded-2xl bg-gradient-to-r from-blue-900/40 via-slate-900/60 to-slate-950 border border-blue-500/10 shadow-2xl relative">
              <span className="px-3 py-1 bg-blue-500/10 text-cyan-400 text-xs font-extrabold rounded-full border border-blue-500/20 tracking-wider">
                ОРГКОМИТЕТ НФЛ / АДМИН-ПАНЕЛЬ
              </span>
              <h2 className="text-2xl md:text-3xl font-extrabold text-white font-display tracking-tight mt-3">
                УПРАВЛЕНИЕ ГОСТЕВЫМИ ЗАЯВКАМИ
              </h2>
              <p className="text-sm text-slate-400 mt-2 max-w-xl">
                Здесь находятся заявки, поданные капитанами в гостевом режиме (без OAuth-авторизации Google). Вы можете согласовать, перепроверить их составы и перенести в официальную Google-таблицу в один клик.
              </p>

              {/* Total Summary Counters */}
              <div className="grid grid-cols-2 gap-4 mt-6 pt-6 border-t border-white/5">
                <div className="bg-slate-950/80 p-4 rounded-xl border border-white/5 flex flex-col justify-between">
                  <span className="text-slate-500 text-[10px] uppercase font-bold tracking-wider">Черновики (не согласовано)</span>
                  <div className="flex items-baseline gap-2 mt-2">
                    <span className="text-2xl md:text-3xl font-black text-yellow-500 font-mono">
                      {submissions.filter(s => !s.synced).length}
                    </span>
                    <span className="text-slate-400 text-xs font-medium ml-1">команд</span>
                  </div>
                </div>
                <div className="bg-slate-950/80 p-4 rounded-xl border border-white/5 flex flex-col justify-between">
                  <span className="text-slate-500 text-[10px] uppercase font-bold tracking-wider">Согласовано (в таблице)</span>
                  <div className="flex items-baseline gap-2 mt-2">
                    <span className="text-2xl md:text-3xl font-black text-emerald-400 font-mono">
                      {submissions.filter(s => s.synced).length}
                    </span>
                    <span className="text-slate-400 text-xs font-medium ml-1">команд</span>
                  </div>
                </div>
              </div>
            </div>

            {/* Deadline Configuration */}
            <div className="bg-slate-950/60 backdrop-blur-md border border-white/5 rounded-2xl p-5 shadow-xl relative mt-6 mb-6">
              <div className="absolute top-0 left-6 w-16 h-[2px] bg-gradient-to-r from-orange-500 to-transparent" />
              <h3 className="text-sm font-bold text-white mb-3 flex items-center gap-2">
                <span className="flex h-2 w-2 relative">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-orange-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-2 w-2 bg-orange-500"></span>
                </span>
                ⏳ Дедлайн (Окончание трансферного окна)
              </h3>
              <p className="text-xs text-slate-400 mb-4">
                Заявки, поданные после этой даты, будут проверяться на наличие новых игроков (по сравнению с прошлыми версиями заявки этой же команды). Новые игроки будут выделены цветом.
              </p>
              
              <div className="bg-slate-900 border border-white/5 rounded-xl p-4 flex items-center justify-between gap-3 max-w-sm">
                <span className="text-sm font-semibold text-slate-300">Дата Дедлайна</span>
                <div className="flex items-center gap-1 shrink-0">
                  <input 
                    type="text"
                    placeholder="ДД.ММ.ГГГГ"
                    value={deadlineDate}
                    onChange={(e) => {
                      const val = e.target.value;
                      const digits = val.replace(/\D/g, '').substring(0, 8);
                      let formatted = '';
                      if (digits.length > 0) formatted += digits.substring(0, 2);
                      if (digits.length > 2) formatted += '.' + digits.substring(2, 4);
                      if (digits.length > 4) formatted += '.' + digits.substring(4, 8);
                      
                      setDeadlineDate(formatted);
                      localStorage.setItem('NFL_DEADLINE_DATE', formatted);
                    }}
                    className="w-[100px] bg-slate-950 border border-white/5 rounded px-2 py-1.5 text-center font-mono text-[13px] text-orange-400 focus:outline-none focus:ring-1 focus:ring-orange-400"
                  />
                </div>
              </div>
            </div>

            {/* Dynamic Tournament Dates Configuration */}
            <div className="bg-slate-950/60 backdrop-blur-md border border-white/5 rounded-2xl p-5 shadow-xl relative">
              <div className="absolute top-0 left-6 w-16 h-[2px] bg-gradient-to-r from-blue-500 to-transparent" />
              <h3 className="text-sm font-bold text-white mb-3 flex items-center gap-2">
                <span className="flex h-2 w-2 relative">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-cyan-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-2 w-2 bg-cyan-500"></span>
                </span>
                📅 Даты проведения турниров для проверки лимита 16 лет
              </h3>
              <p className="text-xs text-slate-400 mb-4">
                Укажите дату первого матча. Система будет автоматически рассчитывать возраст игроков в заявке на основе выбранной даты отборов/финала. Изменения сохраняются моментально.
              </p>
              
              {adminStageTab === 'final' ? (
                <div className="bg-slate-900 border border-white/5 rounded-xl p-4 flex items-center justify-between gap-3 max-w-sm">
                  <span className="text-sm font-semibold text-slate-300">Дата Финального Этапа</span>
                  <div className="flex items-center gap-1 shrink-0">
                    <input 
                      type="text"
                      placeholder="ДД.ММ.ГГГГ"
                      value={finalDate}
                      onChange={(e) => {
                        const val = e.target.value;
                        const digits = val.replace(/\D/g, '').substring(0, 8);
                        let formatted = '';
                        if (digits.length > 0) formatted += digits.substring(0, 2);
                        if (digits.length > 2) formatted += '.' + digits.substring(2, 4);
                        if (digits.length > 4) formatted += '.' + digits.substring(4, 8);
                        
                        setFinalDate(formatted);
                        localStorage.setItem('NFL_FINAL_DATE', formatted);
                      }}
                      className="w-[100px] bg-slate-950 border border-white/5 rounded px-2 py-1.5 text-center font-mono text-[13px] text-[#c5a85c] focus:outline-none focus:ring-1 focus:ring-[#c5a85c]"
                    />
                  </div>
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-h-64 overflow-y-auto pr-1">
                  {dbZones.map(zone => (
                    <div key={zone} className="bg-slate-900 border border-white/5 rounded-xl p-2.5 flex items-center justify-between gap-3">
                      <span className="text-xs font-semibold text-slate-300 truncate" title={zone}>{zone}</span>
                      <div className="flex items-center gap-1 shrink-0">
                        <input 
                          type="text"
                          placeholder="ДД.ММ.ГГГГ"
                          value={zoneDates[zone] || '01.06.2026'}
                          onChange={(e) => {
                            const val = e.target.value;
                            const digits = val.replace(/\D/g, '').substring(0, 8);
                            let formatted = '';
                            if (digits.length > 0) formatted += digits.substring(0, 2);
                            if (digits.length > 2) formatted += '.' + digits.substring(2, 4);
                            if (digits.length > 4) formatted += '.' + digits.substring(4, 8);
                            
                            const nextDates = { ...zoneDates, [zone]: formatted };
                            setZoneDates(nextDates);
                            localStorage.setItem('NFL_ZONE_DATES', JSON.stringify(nextDates));
                          }}
                          className="w-[95px] bg-slate-950 border border-white/5 rounded px-2 py-1 text-center font-mono text-[11px] text-[#c5a85c] focus:outline-none focus:ring-1 focus:ring-[#c5a85c]"
                        />
                      </div>
                    </div>
                  ))}
                  {dbZones.length === 0 && (
                    <div className="col-span-2 text-center py-4 text-xs text-slate-500 italic">
                      Загрузка районов из Google-таблицы (лист "Zones")...
                    </div>
                  )}
                </div>
              )}
            </div>

            {loadingSubmissions ? (
              <div className="flex flex-col items-center justify-center p-12 bg-slate-950/40 rounded-2xl border border-white/5">
                <Loader2 className="w-10 h-10 text-[#c5a85c] animate-spin mb-3" />
                <span className="text-sm text-slate-400 font-medium">Загрузка гостевых заявок из базы...</span>
              </div>
            ) : (
              <div className="space-y-6">
                <div className="flex flex-col gap-4 mb-6">
                  {/* ADMIN STAGE SELECTOR */}
                  <div className="flex bg-slate-900 border border-white/5 rounded-xl p-1 w-full max-w-sm">
                    <button
                      type="button"
                      onClick={() => setAdminStageTab('qualifier')}
                      className={clsx(
                        "flex-1 py-2 text-xs font-bold rounded-lg transition-all",
                        adminStageTab === 'qualifier' ? "bg-slate-800 text-white shadow" : "text-slate-400 hover:text-slate-300"
                      )}
                    >
                      Режим: Отбор
                    </button>
                    <button
                      type="button"
                      onClick={() => setAdminStageTab('final')}
                      className={clsx(
                        "flex-1 py-2 text-xs font-bold rounded-lg transition-all",
                        adminStageTab === 'final' ? "bg-slate-800 text-[#c5a85c] shadow" : "text-slate-400 hover:text-[#c5a85c]/50"
                      )}
                    >
                      Режим: Финал
                    </button>
                  </div>

                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div className="flex bg-slate-900 border border-white/5 rounded-xl p-1 w-full max-w-sm">
                      <button
                        type="button"
                        onClick={() => setAdminTab('drafts')}
                        className={clsx(
                          "flex-1 py-2 text-xs font-bold rounded-lg transition-all",
                          adminTab === 'drafts' ? "bg-slate-800 text-white shadow" : "text-slate-400 hover:text-slate-300"
                        )}
                      >
                        Черновики ({submissions.filter(s => !s.synced && (s.stage || 'qualifier') === adminStageTab).length})
                      </button>
                      <button
                        type="button"
                        onClick={() => setAdminTab('approved')}
                        className={clsx(
                          "flex-1 py-2 text-xs font-bold rounded-lg transition-all",
                          adminTab === 'approved' ? "bg-slate-800 text-green-400 shadow" : "text-slate-400 hover:text-green-400/50"
                        )}
                      >
                        Согласованные ({submissions.filter(s => s.synced && (s.stage || 'qualifier') === adminStageTab).length})
                      </button>
                    </div>

                    <div className="flex flex-wrap gap-2.5">
                    <button
                      type="button"
                      onClick={handleExportAllToExcel}
                      className="flex items-center gap-2 px-4 py-2 bg-emerald-700/80 hover:bg-emerald-600 border border-emerald-500/20 text-white text-xs font-bold rounded-xl transition-all cursor-pointer shadow-lg shadow-emerald-500/10 active:scale-[0.98]"
                    >
                      <FileSpreadsheet className="w-4 h-4 text-emerald-300" />
                      Экспорт всех в Excel
                    </button>
                    <button
                      type="button"
                      onClick={handlePrintAllToPdf}
                      className="flex items-center gap-2 px-4 py-2 bg-[#c5a85c]/80 hover:bg-[#c5a85c] text-white hover:text-slate-950 text-xs font-bold rounded-xl transition-all cursor-pointer shadow-lg shadow-[#c5a85c]/10 active:scale-[0.98]"
                    >
                      <Printer className="w-4 h-4" />
                      Печать всех в PDF
                    </button>
                  </div>
                </div>
                </div>
                
                {(() => {
                   const visibleSubmissions = submissions.filter(s => (adminTab === 'approved' ? s.synced : !s.synced) && (s.stage || 'qualifier') === adminStageTab);
                   if (visibleSubmissions.length === 0) {
                     return (
                       <div className="text-center py-20 bg-slate-950/40 rounded-2xl border border-white/5 border-dashed">
                         <ListTodo className="w-12 h-12 text-slate-600 mx-auto mb-4" />
                         <h3 className="text-lg font-bold text-white mb-2">Пусто</h3>
                         <p className="text-sm text-slate-400 max-w-sm mx-auto leading-relaxed">
                           В этой категории пока нет заявок.
                         </p>
                       </div>
                     );
                   }

                   // Group by zone
                   const grouped: Record<string, typeof submissions> = {};
                   visibleSubmissions.forEach(sub => {
                     const z = sub.zone || 'Неизвестно';
                     if (!grouped[z]) grouped[z] = [];
                     grouped[z].push(sub);
                   });

                   return Object.entries(grouped).map(([zone, subsInZone]) => (
                     <div key={zone} className="border border-white/5 bg-slate-950/60 rounded-xl overflow-hidden mb-4">
                       <button
                         type="button"
                         onClick={() => setExpandedAdminZones(prev => ({...prev, [zone]: !prev[zone]}))}
                         className="w-full flex items-center justify-between p-4 bg-slate-900 border-b border-white/5 text-left hover:bg-slate-800 transition-colors"
                       >
                         <h3 className="text-sm font-bold text-white flex items-center gap-2">
                           {zone}
                           <span className="text-[10px] bg-slate-950 px-2.5 py-0.5 rounded-full text-slate-400 font-mono">
                             {subsInZone.length} {adminTab === 'drafts' ? 'черновиков' : 'согласованных'} (отбор: {zoneDates[zone] || '01.06.2026'})
                           </span>
                         </h3>
                         <ChevronDown className={clsx("w-5 h-5 text-slate-400 transition-transform duration-300", expandedAdminZones[zone] ? "rotate-180" : "")} />
                       </button>
                       <div className={clsx(
                         "flex flex-col divide-y divide-white/5 transition-all overflow-hidden",
                         expandedAdminZones[zone] ? "max-h-none opacity-100" : "max-h-0 opacity-0"
                       )}>
                         {subsInZone.map(sub => (
                           <div key={sub.id} className="p-4 flex flex-col">
                             <button
                               onClick={() => setOpenSubmissionId(openSubmissionId === sub.id ? null : sub.id)}
                               className="flex items-center justify-between w-full text-left"
                             >
                                <div className="flex flex-col md:flex-row md:items-center gap-2">
                                    <div className="flex items-center gap-2">
                                      <span className="font-bold text-white text-sm">{sub.teamName}</span>
                                      {sub.id && (
                                        <div
                                          onClick={(e) => {
                                            e.preventDefault();
                                            e.stopPropagation();
                                            const newName = window.prompt("Изменить название команды (ТОЛЬКО ДЛЯ ЭТОЙ ЗАЯВКИ):", sub.teamName);
                                            if (newName && newName.trim() !== sub.teamName) {
                                              handleUpdateTeamName(sub.id!, newName.trim());
                                            }
                                          }}
                                          className="p-1.5 bg-slate-800/80 hover:bg-slate-700 text-slate-400 hover:text-cyan-400 rounded transition-all cursor-pointer z-10 flex items-center justify-center shrink-0"
                                          title="Переименовать команду"
                                        >
                                          <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 20h9"></path><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"></path></svg>
                                        </div>
                                      )}
                                    </div>
                                  <span className="text-[10px] text-slate-500 font-mono hidden md:inline">•</span>
                                  <span className="text-[10px] text-cyan-400 bg-blue-500/10 border border-blue-500/20 px-2 py-0.5 rounded-full font-bold">Версия {getSubmissionVersion(sub)}</span>
                                  <span className="text-[10px] text-slate-400 bg-slate-900 border border-white/5 px-2 py-0.5 rounded-full font-mono ml-0 md:ml-2">{formatRussianDate(sub.createdAt, true)}</span>
                                </div>
                                <div className="flex items-center gap-3">
                                  {adminTab === 'drafts' ? (
                                     <span className="w-2 h-2 rounded-full bg-yellow-500 animate-pulse"></span>
                                  ) : (
                                     <span className="w-2 h-2 rounded-full bg-green-500"></span>
                                  )}
                                  <ChevronDown className={clsx("w-4 h-4 text-slate-400 transition-transform", openSubmissionId === sub.id ? "rotate-180" : "")} />
                                </div>
                             </button>
                             
                             {openSubmissionId === sub.id && (
                                <div className="mt-4 bg-slate-900/50 p-4 rounded-xl border border-white/5">
                                  <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-4 border-b border-white/5 pb-4">
                                     <div className="text-xs text-slate-400 leading-relaxed">
                                       <p><span className="font-bold text-slate-300 uppercase tracking-wider text-[10px]">Капитан:</span> <span className="text-white ml-1">{sub.captainName}</span></p>
                                       <p className="mt-1"><span className="font-bold text-slate-300 uppercase tracking-wider text-[10px]">Телефон:</span> <a href={`tel:${sub.captainPhone}`} className="text-cyan-400 ml-1 hover:underline">{sub.captainPhone}</a></p>
                                     </div>
                                     <div className="flex flex-wrap gap-2 w-full md:w-auto">
                                       <button
                                         type="button"
                                         onClick={() => handleExportTeamToExcel(sub)}
                                         className="flex-1 md:flex-none justify-center bg-slate-800 hover:bg-slate-700 text-emerald-400 border border-emerald-300/10 text-xs font-bold px-3.5 py-2 rounded-lg flex items-center gap-1.5 transition-all active:scale-[0.97] cursor-pointer"
                                         title="Скачать состав в Excel"
                                       >
                                         <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-300" />
                                         Excel
                                       </button>
                                       <button
                                         type="button"
                                         onClick={() => handlePrintTeamToPdf(sub)}
                                         className="flex-1 md:flex-none justify-center bg-slate-800 hover:bg-slate-700 text-[#c5a85c] border border-[#c5a85c]/10 text-xs font-bold px-3.5 py-2 rounded-lg flex items-center gap-1.5 transition-all active:scale-[0.97] cursor-pointer"
                                         title="Печать состава в PDF"
                                       >
                                         <Printer className="w-3.5 h-3.5" />
                                         PDF
                                       </button>
                                       {adminTab === 'drafts' ? (
                                         <>
                                           <button
                                             onClick={() => handleSyncToSheets(sub, 'Согласованные')}
                                             disabled={syncingSubmissionId !== null}
                                             className="flex-1 md:flex-none justify-center bg-gradient-to-r from-[#c5a85c] to-[#e4cb8c] hover:scale-[1.02] active:scale-[0.98] text-slate-950 text-xs font-bold px-4 py-2 rounded-lg flex items-center gap-1.5 transition-all shadow-lg disabled:opacity-50"
                                           >
                                             {syncingSubmissionId === sub.id ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Check className="w-3.5 h-3.5" />} Согласовать
                                           </button>
                                           <button
                                             onClick={() => handleDeleteSubmission(sub)}
                                             className="flex-1 md:flex-none justify-center bg-slate-900 border border-slate-700/50 hover:bg-red-500/10 hover:border-red-500/20 hover:text-red-400 text-slate-400 text-xs font-bold px-3 py-2 rounded-lg flex items-center gap-1.5 transition-all"
                                             title="Удалить заявку"
                                           >
                                             <Trash2 className="w-3.5 h-3.5" />
                                           </button>
                                         </>
                                       ) : (
                                         <button
                                           onClick={() => handleRevokeApproval(sub)}
                                           disabled={syncingSubmissionId !== null}
                                           className="flex-1 md:flex-none justify-center bg-slate-900 text-red-400 border border-red-500/20 hover:bg-red-950/80 text-xs font-bold px-4 py-2 rounded-lg flex items-center gap-1.5 transition-all disabled:opacity-50"
                                         >
                                           {syncingSubmissionId === sub.id ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <X className="w-3.5 h-3.5" />} Отозвать
                                         </button>
                                       )}
                                     </div>
                                  </div>
                                  
                                  {/* Roster display */}
                                  <div className="space-y-1.5">
                                    <h4 className="flex items-center gap-2 text-[10px] font-bold tracking-wider text-slate-500 mb-3 uppercase">
                                      <Users className="w-3.5 h-3.5" /> Состав команды:
                                    </h4>
                                    {sub.players.map((p, idx) => {
                                      const duplicateInTeam = getPlayerApprovedTeam(p.fullName, sub.teamName);
                                      const isAfterDeadlineMatch = deadlineDate?.length === 10 ? new Date(sub.createdAt || 0).getTime() > parseDDMMYYYYToTime(deadlineDate) : false;
                                      const isNewAddition = isPlayerNewlyAdded(sub, p.fullName);
                                      const requiresFee = isAfterDeadlineMatch && isNewAddition;
                                      
                                      return (
                                        <div 
                                          key={p.id || idx} 
                                          className={clsx(
                                            "grid grid-cols-[16px_1fr_60px_40px_35px] md:grid-cols-[20px_1fr_80px_60px_40px] items-center text-[10px] md:text-xs p-1.5 rounded-lg transition-colors border",
                                            duplicateInTeam 
                                              ? "bg-red-950/40 border-red-500/20 text-red-200 shadow-[0_0_15px_rgba(239,68,68,0.15)]" 
                                              : requiresFee
                                                ? "bg-orange-950/40 border-orange-500/20 text-orange-200 shadow-[0_0_15px_rgba(249,115,22,0.15)]"
                                                : "hover:bg-white/5 border-transparent text-slate-300"
                                          )}
                                        >
                                          <span className={clsx("font-mono", requiresFee ? "text-orange-400" : "text-slate-600")}>{idx + 1}.</span>
                                          <span className={clsx(
                                            "font-semibold truncate pr-2 flex flex-wrap items-center gap-1.5",
                                            duplicateInTeam ? "text-red-300 font-extrabold" : (requiresFee ? "text-orange-300 font-extrabold" : "text-slate-300")
                                          )}>
                                            {p.fullName || 'Без имени'}
                                            {p.isLegionnaire && (
                                              <span className="text-red-400 border border-red-500/30 bg-red-500/10 px-1 rounded-sm text-[8px] font-bold" title="Легионер">
                                                ЛЕГ
                                              </span>
                                            )}
                                            {requiresFee && (
                                              <span className="bg-orange-600/30 text-orange-400 border border-orange-500/40 text-[9px] px-1.5 py-0.5 rounded font-bold uppercase tracking-wider ml-1 whitespace-nowrap" title="Изменение внесено после дедлайна - может требоваться комиссия">АПДЕЙТ+</span>
                                            )}
                                            {p.transferStatus === 'new' && (
                                              <span className="text-cyan-400 border border-cyan-500/30 bg-cyan-500/10 px-1 rounded-sm text-[8px] font-bold" title="Новый игрок">
                                                НОВ.
                                              </span>
                                            )}
                                            {p.transferStatus === 'other_zone' && (
                                              <span className="text-amber-400 border border-amber-500/30 bg-amber-500/10 px-1 rounded-sm text-[8px] font-bold" title="Заигран в другой зоне">
                                                ДР.З.
                                              </span>
                                            )}
                                            {p.isVerified ? (
                                              <span className="w-1.5 h-1.5 bg-green-500 rounded-full" title="Заигран"></span>
                                            ) : (
                                              <React.Fragment>
                                                {p.fullName && dbPlayers.some(db => db.fullName.toLowerCase() === p.fullName.toLowerCase()) && (
                                                  <FileText className="w-3 h-3 text-orange-400" title="Был в базе, но не заигран" />
                                                )}
                                              </React.Fragment>
                                            )}
                                            {duplicateInTeam && (
                                              <span 
                                                className="text-[8px] bg-red-600 hover:bg-red-500 text-white font-extrabold px-1.5 py-0.5 rounded animate-pulse text-[8px] shrink-0" 
                                                title={`Игрок уже согласован в составе другой команды: "${duplicateInTeam}"!`}
                                              >
                                                ⚠️ КОПИЯ В {duplicateInTeam.toUpperCase()}
                                              </span>
                                            )}
                                          </span>
                                          <span className="text-slate-500 font-mono">{p.birthDate || '-'}</span>
                                          <span className="text-slate-400">{p.position || '-'}</span>
                                          <span className="text-slate-400 text-center font-mono">#{p.number || '-'}</span>
                                        </div>
                                      );
                                    })}
                                  </div>
                                </div>
                             )}
                           </div>
                         ))}
                       </div>
                     </div>
                   ));
                })()}
              </div>
            )}

          </div>
        ) : !stage ? (
          <div className="flex flex-col items-center justify-center py-20 px-4 min-h-[50vh]">
            <div className="bg-slate-900 border border-slate-700 p-8 rounded-2xl max-w-lg w-full text-center shadow-2xl relative overflow-hidden">
               <div className="absolute top-0 right-0 w-48 h-full bg-[radial-gradient(circle_at_right,_var(--tw-gradient-stops))] from-[#c5a85c]/10 to-transparent pointer-events-none rounded-r-2xl" />
               <Trophy className="w-16 h-16 text-[#c5a85c] mx-auto mb-6" />
               <h2 className="text-2xl font-extrabold text-white mb-2 font-display">Выберите этап турнира</h2>
               <p className="text-slate-400 text-sm mb-8">Для правильного формирования заявки выберите этап, на который подается состав команды.</p>
               <div className="grid grid-cols-1 md:grid-cols-2 gap-4 w-full">
                 <button 
                   onClick={() => setStage('qualifier')}
                   className="flex flex-col items-center justify-center bg-slate-950 border border-slate-700 hover:border-blue-500 hover:bg-blue-900/20 text-white p-6 rounded-xl transition-all cursor-pointer group shadow-lg"
                 >
                   <span className="text-lg font-bold group-hover:text-cyan-400 transition-colors">Отборочный Этап</span>
                   <span className="text-xs text-slate-500 mt-2">Формирование базового состава</span>
                 </button>
                 <button 
                   onClick={() => setStage('final')}
                   className="flex flex-col items-center justify-center bg-slate-950 border border-slate-700 hover:border-[#c5a85c] hover:bg-[#c5a85c]/10 text-white p-6 rounded-xl transition-all cursor-pointer group shadow-lg"
                 >
                   <span className="text-lg font-bold group-hover:text-[#c5a85c] transition-colors">Финальный Этап</span>
                   <span className="text-xs text-slate-500 mt-2">Дозаявки и изменения статусов</span>
                 </button>
               </div>
            </div>
          </div>
        ) : (
          <>
            {/* League Champions Banner Header */}
            <div className="mb-8 p-6 md:p-8 rounded-2xl bg-gradient-to-r from-blue-900/40 via-slate-900/60 to-slate-950 border border-blue-500/10 flex flex-col md:flex-row items-center justify-between gap-6 shadow-[0_10px_40px_rgba(0,0,0,0.5)] relative">
              <div className="absolute top-0 right-0 w-32 h-full bg-[radial-gradient(circle_at_right,_var(--tw-gradient-stops))] from-[#c5a85c]/5 to-transparent pointer-events-none rounded-r-2xl" />
              
              <div className="text-center md:text-left">
                <span className="px-3 py-1 bg-blue-500/10 text-cyan-400 text-xs font-extrabold rounded-full border border-blue-500/20 tracking-wider">
                  СЕЗОН 2026 / ОТБОРОЧНЫЙ ЭТАП
                </span>
                <h2 className="text-2xl md:text-3xl font-extrabold text-white font-display tracking-tight mt-3">
                  ЗАЯВКА НА ОТБОР В НОГАЙСКУЮ ФУТБОЛЬНУЮ ЛИГУ
                </h2>
                <p className="text-sm text-slate-400 mt-2 max-w-xl">
                  Заполните анкету своей футбольной команды. Вы можете загрузить состав за прошлые годы, отредактировать игроков, или внести новые имена.
                </p>
              </div>
              <div className="hidden md:flex bg-slate-950/85 border border-white/5 p-5 rounded-2xl items-center justify-center text-[#c5a85c] shadow-[0_0_20px_rgba(197,168,92,0.1)]">
                <Trophy className="w-10 h-10" />
              </div>
            </div>

            {/* Instruction Infographic */}
            <div className="mb-8 p-5 md:p-6 bg-slate-900/40 border border-[#c5a85c]/10 rounded-2xl relative overflow-hidden">
              <div className="absolute top-0 right-0 w-32 h-32 bg-gradient-to-br from-[#c5a85c]/5 to-transparent rounded-full filter blur-[40px] pointer-events-none" />
              
              <h3 className="text-sm md:text-base font-bold text-white mb-6 flex items-center gap-2">
                <Info className="w-5 h-5 text-[#c5a85c]" />
                Как заполнить заявку? 
              </h3>
              
              <div className="flex flex-col gap-5 relative">
                {/* Vertical Line */}
                <div className="absolute left-[15px] top-4 bottom-4 w-[2px] bg-gradient-to-b from-slate-800 via-slate-800 to-[#c5a85c]/30" />
                
                {/* Step 1 */}
                <div className="flex items-start gap-4 relative z-10">
                  <div className="w-8 h-8 rounded-full bg-slate-950 border border-slate-700 flex items-center justify-center text-xs font-bold text-slate-300 shrink-0 shadow-lg">1</div>
                  <div className="pt-1.5">
                    <p className="text-sm font-bold text-slate-200">Впишите данные команды</p>
                    <p className="text-xs text-slate-400 mt-1">Начните вводить название — и мы <strong className="text-slate-300">автоматически подгрузим</strong> игроков, которые играли за вас ранее.</p>
                  </div>
                </div>
                
                {/* Step 2 */}
                <div className="flex items-start gap-4 relative z-10">
                  <div className="w-8 h-8 rounded-full bg-slate-950 border border-slate-700 flex items-center justify-center text-xs font-bold text-slate-300 shrink-0 shadow-lg">2</div>
                  <div className="pt-1.5">
                    <p className="text-sm font-bold text-slate-200">Отредактируйте состав</p>
                    <p className="text-xs text-slate-400 mt-1">Оставьте тех, кто играет в этом сезоне. Удалите лишних крестиком <X className="inline w-3 h-3 text-red-400 mb-0.5" /> справа.</p>
                  </div>
                </div>

                {/* Step 3 */}
                <div className="flex items-start gap-4 relative z-10">
                  <div className="w-8 h-8 rounded-full bg-slate-950 border border-slate-700 flex items-center justify-center text-xs font-bold text-slate-300 shrink-0 shadow-lg">3</div>
                  <div className="pt-1.5">
                    <p className="text-sm font-bold text-slate-200">Добавьте новичков</p>
                    <p className="text-xs text-slate-400 mt-1">Нажмите «Добавить футболиста» внизу списка. Помните: <strong className="text-red-400">максимум 1 легионер</strong>.</p>
                  </div>
                </div>

                {/* Step 4 */}
                <div className="flex items-start gap-4 relative z-10">
                  <div className="w-8 h-8 rounded-full bg-slate-950 border border-[#c5a85c]/50 flex items-center justify-center text-xs font-bold text-[#c5a85c] shrink-0 shadow-[0_0_10px_rgba(197,168,92,0.2)]">
                    <Check className="w-4 h-4" />
                  </div>
                  <div className="pt-1.5">
                    <p className="text-sm font-bold text-[#c5a85c]">Отправьте заявку</p>
                    <p className="text-xs text-[#c5a85c]/70 mt-1">Готово! Заявка уйдет на проверку оргкомитету НФЛ.</p>
                  </div>
                </div>
              </div>
            </div>

            <form onSubmit={handleSubmit} className="space-y-8">

          {/* Section: Основная информация */}
          <div className="bg-slate-950/60 backdrop-blur-md border border-white/5 rounded-2xl p-6 shadow-2xl relative">
             <div className="absolute top-0 left-6 w-16 h-[2px] bg-gradient-to-r from-[#c5a85c] to-transparent" />
             
             <h2 className="text-lg font-bold font-display text-white mb-6 flex items-center gap-2.5 border-b border-white/5 pb-3">
               <span className="bg-[#c5a85c] text-slate-950 w-7 h-7 rounded-full flex items-center justify-center text-xs font-extrabold shadow-[0_0_15px_rgba(197,168,92,0.4)]">1</span>
               Основная информация о команде
             </h2>

             <div className="space-y-4">
                 <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="relative">
                        <label className="block text-xs font-bold uppercase tracking-wider text-slate-400 mb-1.5">Название команды</label>
                        <input 
                          type="text" 
                          required
                          placeholder="Введите название..."
                          value={form.teamName}
                          onChange={handleTeamNameChange}
                          onFocus={() => setShowTeamSuggestions(true)}
                          onBlur={() => setTimeout(() => setShowTeamSuggestions(false), 200)}
                          className="w-full bg-slate-900/80 border border-white/5 rounded-xl px-4 py-3 text-sm placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-[#c5a85c] focus:border-transparent transition-all"
                        />
                        {showTeamSuggestions && (
                          <div className="absolute left-0 right-0 top-full mt-1 bg-slate-950 border border-white/10 rounded-xl shadow-[0_10px_30px_rgba(0,0,0,0.5)] z-50 overflow-hidden max-h-60 overflow-y-auto">
                            {suggestedTeams
                              .filter(t => !form.teamName || t.trim().toLowerCase().includes(form.teamName.trim().toLowerCase()))
                              .map(t => (
                                <div
                                  key={t}
                                  onMouseDown={(e) => {
                                    e.preventDefault();
                                    handleTeamSelect(t);
                                    setShowTeamSuggestions(false);
                                  }}
                                  className="px-4 py-2.5 hover:bg-white/5 text-sm cursor-pointer border-b border-white/5 last:border-b-0 text-slate-300 hover:text-white transition-colors flex items-center justify-between"
                                >
                                  <span>{t}</span>
                                  <span className="text-[10px] bg-blue-500/10 text-cyan-400 font-mono px-1.5 py-0.5 rounded uppercase tracking-wider font-bold">Команда</span>
                                </div>
                              ))
                            }
                            {suggestedTeams.filter(t => !form.teamName || t.trim().toLowerCase().includes(form.teamName.trim().toLowerCase())).length === 0 && (
                              <div className="px-4 py-3 text-xs text-slate-500 italic">Команда не найдена (будет создана новая)</div>
                            )}
                          </div>
                        )}
                        {suggestedTeams.find(t => t.trim().toLowerCase() === form.teamName.trim().toLowerCase()) && (
                          <button 
                             id="load-prev-squad-btn"
                             type="button"
                             onClick={() => {
                               if (form.players.length === 0 || window.confirm("Вы уверены, что хотите загрузить состав команды из базы данных? Текущие внесенные игроки будут заменены.")) {
                                 handleTeamSelect(form.teamName);
                               }
                             }}
                             className="mt-2 text-xs text-[#c5a85c] hover:text-white font-bold flex items-center gap-1.5 transition-colors cursor-pointer animate-pulse"
                          >
                            ⚡ Загрузить предыдущий состав команды из базы {form.players.length > 0 && "(перезаписать текущий)"}
                          </button>
                        )}
                     </div>

                     <div className="relative">
                        <label className="block text-xs font-bold uppercase tracking-wider text-slate-400 mb-1.5">Район / Зона</label>
                        <input 
                          type="text" 
                          required
                          placeholder="Выберите или введите..."
                          value={form.zone}
                          onChange={(e) => setForm(prev => ({ ...prev, zone: e.target.value }))}
                          onFocus={() => setShowZoneSuggestions(true)}
                          onBlur={() => setTimeout(() => setShowZoneSuggestions(false), 200)}
                          className="w-full bg-slate-900/80 border border-white/5 rounded-xl px-4 py-3 text-sm placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-[#c5a85c] focus:border-transparent transition-all"
                        />
                        {showZoneSuggestions && (
                          <div className="absolute left-0 right-0 top-full mt-1 bg-slate-950 border border-white/10 rounded-xl shadow-[0_10px_30px_rgba(0,0,0,0.5)] z-50 overflow-hidden max-h-60 overflow-y-auto">
                            {dbZones
                              .filter(z => !form.zone || z.toLowerCase().includes(form.zone.toLowerCase()))
                              .map(z => (
                                <div
                                  key={z}
                                  onMouseDown={(e) => {
                                    e.preventDefault();
                                    setForm(prev => ({ ...prev, zone: z }));
                                    setShowZoneSuggestions(false);
                                  }}
                                  className="px-4 py-2.5 hover:bg-white/5 text-sm cursor-pointer border-b border-white/5 last:border-b-0 text-slate-300 hover:text-white transition-colors flex items-center justify-between"
                                >
                                  <span>{z}</span>
                                  <span className="text-[10px] bg-emerald-500/10 text-emerald-400 font-mono px-1.5 py-0.5 rounded uppercase tracking-wider font-bold">Район</span>
                                </div>
                              ))
                            }
                          </div>
                        )}
                     </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-4 border-t border-white/5">
                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-400 mb-1.5">ФИО Капитана</label>
                      <input 
                        type="text" 
                        required
                        placeholder="Капитан команды"
                        value={form.captainName}
                        onChange={(e) => setForm(prev => ({...prev, captainName: e.target.value}))}
                        className="w-full bg-slate-900/80 border border-white/5 rounded-xl px-4 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-[#c5a85c] focus:border-transparent transition-all"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold uppercase tracking-wider text-slate-400 mb-1.5">Телефон Капитана</label>
                      <input 
                        type="text" 
                        required
                        placeholder="+7 (___) ___-__-__"
                        value={form.captainPhone}
                        onChange={(e) => setForm(prev => ({...prev, captainPhone: e.target.value}))}
                        className="w-full bg-slate-900/80 border border-white/5 rounded-xl px-4 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-[#c5a85c] focus:border-transparent transition-all"
                      />
                    </div>
                  </div>
              </div>
          </div>

          {/* Section: Состав команды */}
          <div className="bg-slate-950/60 backdrop-blur-md border border-white/5 rounded-2xl p-6 shadow-2xl relative">
             <div className="absolute top-0 left-6 w-16 h-[2px] bg-gradient-to-r from-blue-500 to-transparent" />
             
             <div className="flex items-center justify-between mb-6 border-b border-white/5 pb-4">
                <h2 className="text-lg font-bold font-display text-white flex items-center gap-2.5">
                  <span className="bg-gradient-to-tr from-blue-700 to-blue-500 text-white w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold shadow-[0_0_15px_rgba(37,99,235,0.4)]">2</span>
                  Заявка игроков в команду
                </h2>
             </div>

             {/* Dynamic Player count indicator bar / regulators */}
             <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-slate-900/50 p-3 rounded-xl border border-white/5 mb-6 text-xs text-slate-300">
               <div className="flex items-center gap-2">
                 <Users className="w-4 h-4 text-[#c5a85c]" />
                 <span>Лимит заявки: от 10 до 15 футболистов</span>
               </div>
               <div className="flex items-center gap-2">
                 <span className="font-semibold text-slate-400">Вбито игроков:</span>
                 <span className={clsx(
                   "font-mono font-bold px-2 py-0.5 rounded-md border text-sm",
                   form.players.filter(p => p.status !== 'deleted' && p.fullName.trim() !== '').length >= 10 && form.players.filter(p => p.status !== 'deleted' && p.fullName.trim() !== '').length <= 15
                     ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/20 shadow-[0_0_10px_rgba(16,185,129,0.1)]"
                     : "bg-amber-500/10 text-amber-400 border-amber-500/20 shadow-[0_0_10px_rgba(245,158,11,0.1)]"
                 )}>
                   {form.players.filter(p => p.status !== 'deleted' && p.fullName.trim() !== '').length} из 15
                 </span>
                 {form.players.filter(p => p.status !== 'deleted' && p.fullName.trim() !== '').length < 10 && (
                   <span className="text-red-400 font-bold font-mono text-[11px] animate-pulse">
                     (нужно ещё {10 - form.players.filter(p => p.status !== 'deleted' && p.fullName.trim() !== '').length})
                   </span>
                 )}
               </div>
             </div>

             <div className="space-y-3.5">
                {form.players.length === 0 ? (
                  <div className="text-center py-12 bg-slate-900/40 rounded-xl border border-white/5 border-dashed">
                    <p className="text-slate-400 text-sm max-w-md mx-auto leading-relaxed">
                      В составе пока нет добавленных игроков. Начните вводить название команды выше, чтобы подгрузить прошлый архив, или добавьте новых футболистов вручную.
                    </p>
                  </div>
                ) : (
                  form.players.map((player, index) => {
                    if (player.status === 'deleted') {
                      return (
                        <div key={player.id} className="flex items-center justify-between p-3 bg-red-950/20 border border-red-500/20 rounded-xl opacity-60 text-xs">
                          <div className="flex items-center gap-3">
                            <span className="text-xs font-mono text-slate-500">#{index + 1}</span>
                            <span className="line-through text-slate-400 font-medium">{player.fullName || 'Безымянный игрок'}</span>
                            <span className="text-[10px] uppercase font-bold text-red-500 bg-red-500/10 px-1.5 py-0.5 rounded">Удален</span>
                          </div>
                          <button 
                            type="button" 
                            onClick={() => restorePlayer(player.id)}
                            className="text-xs text-[#c5a85c] hover:underline font-bold cursor-pointer"
                          >
                            Восстановить
                          </button>
                        </div>
                      );
                    }

                    const isPrevious = player.status === 'previous';

                    return (
                      <div 
                        key={player.id} 
                        id={player.id}
                        className={clsx(
                          "p-4 rounded-xl border transition-all duration-300 relative group text-xs",
                          isPrevious 
                            ? "bg-slate-900/60 border-white/5 border-l-4 border-l-yellow-500 shadow-[0_4px_12px_rgba(234,179,8,0.03)]" 
                            : "bg-slate-900/60 border-white/5"
                        )}
                      >
                         {/* Delete button (stands on right side always visible to allow quick deleting) */}
                         <button 
                           type="button" 
                           onClick={() => removePlayer(player.id, player.status)} 
                           className="absolute top-3.5 right-4 text-slate-400 hover:text-red-400 focus:text-red-400 transition-colors p-1 rounded-full hover:bg-white/5" 
                           title="Удалить футболиста"
                         >
                           <X className="w-4 h-4" />
                         </button>

                         <div className="pr-6 md:pr-10 flex flex-col gap-3">
                           {/* Counter Indicator */}
                           <div className="flex items-center gap-1.5 text-[10px] font-bold uppercase tracking-wider text-slate-400">
                             <span className={clsx(
                                "w-4.5 h-4.5 rounded flex items-center justify-center font-mono text-[9px]",
                                isPrevious ? "bg-yellow-500/20 text-yellow-300 border border-yellow-500/30" : "bg-slate-800 text-slate-300"
                             )}>
                               {index + 1}
                             </span>
                             {isPrevious ? "Игрок архивного состава (из базы данных)" : "Новый футболист"}
                           </div>

                           <div className="flex flex-col md:flex-row md:items-start gap-3 w-full">
                             {/* Name, full width on mobile */}
                             <div className="w-full md:flex-1 relative pt-1">
                               <label className="flex flex-wrap items-center justify-between text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-0.5">
                                 <span>ФИО Игрока {player.isLegionnaire && <span className="text-red-400 ml-1">(Легионер)</span>}</span>
                                 {isPrevious && (
                                   <span className="px-1.5 py-0.5 rounded flex items-center justify-center font-mono text-[9px] bg-yellow-500/20 text-yellow-300 border border-yellow-500/30">
                                     АРХИВ #{index + 1}
                                   </span>
                                 )}
                               </label>
                               <div className="relative">
                                 <input 
                                   type="text"
                                   required
                                   placeholder="Имя Фамилия"
                                   value={player.fullName}
                                   onChange={(e) => updatePlayer(player.id, 'fullName', e.target.value)}
                                   onFocus={() => setFocusedPlayerInputId(player.id)}
                                   onBlur={() => setTimeout(() => setFocusedPlayerInputId(null), 200)}
                                   className="w-full bg-slate-950 border border-white/5 rounded-lg px-3 py-1.5 text-xs text-slate-200 focus:outline-none focus:ring-2 focus:ring-[#c5a85c] focus:border-transparent transition-all pr-8"
                                 />
                                 {player.isVerified && isGuest && (
                                   <div className="absolute right-2 top-1/2 -translate-y-1/2 flex items-center justify-center group/doc" title="Заигран (документы проверены)">
                                     <div className="w-4 h-4 bg-green-500/10 rounded flex items-center justify-center border border-green-500/30">
                                       <svg xmlns="http://www.w3.org/2000/svg" width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" className="text-green-400">
                                         <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                         <polyline points="14 2 14 8 20 8"></polyline>
                                         <line x1="16" y1="13" x2="8" y2="13"></line>
                                         <line x1="16" y1="17" x2="8" y2="17"></line>
                                         <polyline points="10 9 9 9 8 9"></polyline>
                                       </svg>
                                     </div>
                                   </div>
                                 )}
                               </div>
                               {focusedPlayerInputId === player.id && getPlayerSuggestions(player.fullName).length > 0 && (
                                 <div className="absolute left-0 right-0 top-full mt-1 bg-slate-950 border border-white/10 rounded-xl shadow-[0_10px_30px_rgba(0,0,0,0.5)] z-50 overflow-hidden max-h-60 overflow-y-auto">
                                   {getPlayerSuggestions(player.fullName).map((s) => (
                                     <div
                                       key={s.fullName}
                                       onMouseDown={(e) => {
                                         e.preventDefault();
                                         updatePlayer(player.id, 'fullName', s.fullName);
                                         if (s.birthDate) updatePlayer(player.id, 'birthDate', s.birthDate);
                                         if (s.position) updatePlayer(player.id, 'position', s.position);
                                         if (s.number) updatePlayer(player.id, 'number', s.number);
                                         if (s.isLegionnaire !== undefined) updatePlayer(player.id, 'isLegionnaire', s.isLegionnaire);
                                         if (s.isVerified !== undefined) updatePlayer(player.id, 'isVerified', s.isVerified);
                                         setFocusedPlayerInputId(null);
                                       }}
                                       className="px-3 py-2 hover:bg-white/5 text-xs cursor-pointer border-b border-white/5 last:border-b-0 text-slate-300 hover:text-white transition-colors flex flex-col gap-0.5 text-left"
                                     >
                                       <span className="font-bold text-white text-xs flex items-center gap-1.5">
                                         {s.fullName}
                                         {s.isVerified ? <span className="w-1.5 h-1.5 bg-green-500 rounded-full" title="Заигран"></span> : <FileText className="w-3 h-3 text-orange-400" title="Был в базе, но не заигран" />}
                                       </span>
                                       <span className="text-[10px] text-slate-400 font-mono">
                                         {s.birthDate || 'ДД.ММ.ГГГГ'} • {s.position || 'Позиция'} • Команда: {s.teamName}
                                       </span>
                                     </div>
                                   ))}
                                 </div>
                               )}
                             </div>
                           {/* Second row on mobile: Date, Position, Number, Legionnaire */}
                           <div className="flex flex-row items-end md:items-start gap-2 w-full md:w-auto shrink-0 pt-1">
                             <div className="flex-1 min-w-[100px] md:w-[120px] relative">
                               <label className="block text-[9px] font-bold uppercase tracking-wider text-slate-500 mb-0.5">Д.Рожд.</label>
                               <input 
                                 type="text"
                                 placeholder="ДД.ММ.ГГГГ"
                                 required
                                 value={player.birthDate}
                                 onChange={(e) => {
                                   let val = e.target.value;
                                   const digits = val.replace(/\D/g, '').substring(0, 8);
                                   let formatted = '';
                                   if (digits.length > 0) formatted += digits.substring(0, 2);
                                   if (digits.length > 2) formatted += '.' + digits.substring(2, 4);
                                   if (digits.length > 4) formatted += '.' + digits.substring(4, 8);
                                   updatePlayer(player.id, 'birthDate', formatted);
                                 }}
                                 className="w-full bg-slate-950 border border-white/5 rounded-lg px-2 py-1.5 text-xs text-slate-200 focus:outline-none focus:ring-2 focus:ring-[#c5a85c] focus:border-transparent transition-all"
                               />
                               {(() => {
                                 const tDate = stage === 'final' ? finalDate : (zoneDates[form.zone] || '01.06.2026');
                                 const playerAge = player.birthDate.length === 10 ? calculateAge(player.birthDate, tDate) : null;
                                 if (playerAge !== null && playerAge < 18 && playerAge >= 16) {
                                   const consentText = `СОГЛАСИЕ РОДИТЕЛЕЙ (ЗАКОННЫХ ПРЕДСТАВИТЕЛЕЙ)
НА УЧАСТИЕ НЕСОВЕРШЕННОЛЕТНЕГО В СОРЕВНОВАНИЯХ ПО ФУТБОЛУ

Я, ____________________________________________________________________,
паспорт (серия, номер, выдан): ________________________________________
____________________________________________________________________,
телефон: _________________________________,
являясь законным представителем несовершеннолетнего:
ФИО ребенка: ${player.fullName || '___________________________________________________________'},
Дата рождения ребенка: ${player.birthDate || '______________________'},

добровольно даю свое согласие на участие моего ребенка в матчах и играх Ногайской Футбольной Лиги (Сезон 2026).

Я подтверждаю, что:
1. Мой ребенок не имеет медицинских противопоказаний для занятий футболом и участия в спортивных соревнованиях.
2. Я полностью осведомлен(а) о возможных рисках получения травм, связанных с данным видом спорта, и принимаю всю ответственность на себя.

Законный представитель:
_________________ / _________________________________ (Подпись / ФИО)

Дата: «_____» _______________ 2026 г.`;
                                   return (
                                     <a 
                                       href={`data:text/plain;charset=utf-8,${encodeURIComponent(consentText)}`} 
                                       download={`Soglasie_${player.fullName || 'Igrok'}.txt`} 
                                       className="absolute -bottom-4 left-0 text-[8px] text-yellow-500 hover:text-yellow-400 hover:underline flex items-center gap-1 font-semibold"
                                     >
                                       💾 Скачать согласие родителей (16-17 лет)
                                     </a>
                                   );
                                 }
                                 return null;
                               })()}
                             </div>
                             <div className="w-[85px] shrink-0">
                               <label className="block text-[9px] font-bold uppercase tracking-wider text-slate-500 mb-0.5">Амплуа</label>
                               <select
                                 required
                                 value={player.position || ""}
                                 onChange={(e) => updatePlayer(player.id, 'position', e.target.value)}
                                 className="w-full bg-slate-950 border border-white/5 text-slate-200 rounded-lg px-1 py-1.5 text-[11px] focus:outline-none focus:ring-2 focus:ring-[#c5a85c] focus:border-transparent transition-all"
                               >
                                 <option value="" disabled className="text-slate-500 bg-slate-950">-</option>
                                 <option value="ВРТ" className="text-slate-200 bg-slate-950">ВРТ</option>
                                 <option value="ЗАЩ" className="text-slate-200 bg-slate-950">ЗАЩ</option>
                                 <option value="ПЗЩ" className="text-slate-200 bg-slate-950">ПЗЩ</option>
                                 <option value="НАП" className="text-slate-200 bg-slate-950">НАП</option>
                               </select>
                             </div>
                             <div className="w-[45px] shrink-0">
                               <label className="block text-[9px] font-bold uppercase tracking-wider text-slate-500 mb-0.5 text-center">№</label>
                               <input 
                                 type="text"
                                 placeholder="№"
                                 value={player.number}
                                 onChange={(e) => updatePlayer(player.id, 'number', e.target.value)}
                                 className="w-full bg-slate-950 border border-white/5 rounded-lg px-1 py-1.5 text-xs text-center text-slate-200 focus:outline-none focus:ring-2 focus:ring-[#c5a85c] focus:border-transparent transition-all"
                               />
                             </div>
                             
                             <div className="flex flex-col items-center justify-center self-stretch shrink-0">
                               <label className="block text-[8px] font-bold uppercase tracking-wider text-slate-500 mb-[3px]">Лег.</label>
                               <button
                                 type="button"
                                 onClick={() => updatePlayer(player.id, 'isLegionnaire', !player.isLegionnaire)}
                                 className={clsx(
                                   "w-6 h-6 rounded flex items-center justify-center transition-colors border",
                                   player.isLegionnaire ? "bg-red-500/20 text-red-400 border-red-500/50" : "bg-slate-900 border-white/10 hover:border-white/30 text-slate-600"
                                 )}
                                 title="Отметить как легионера"
                               >
                                 {player.isLegionnaire ? (
                                   <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
                                     <polyline points="20 6 9 17 4 12"></polyline>
                                   </svg>
                                 ) : (
                                   <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                     <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path>
                                   </svg>
                                 )}
                               </button>
                             </div>
                           </div>
                           
                           {/* Transfer Status Field for Final stage */}
                           {stage === 'final' && (
                             <div className="w-full mt-2 pt-2 border-t border-white/5">
                               <label className="block text-[9px] font-bold uppercase tracking-wider text-slate-500 mb-1">
                                 Статус игрока относительно отборочного этапа
                               </label>
                               <select
                                 required
                                 value={player.transferStatus || ""}
                                 onChange={(e) => updatePlayer(player.id, 'transferStatus', e.target.value as Player['transferStatus'])}
                                 className={clsx(
                                   "w-full border rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:border-transparent transition-all",
                                   player.transferStatus === 'new' ? "bg-cyan-500/10 text-cyan-400 border-cyan-500/30 focus:ring-cyan-500" :
                                   player.transferStatus === 'other_zone' ? "bg-amber-500/10 text-amber-400 border-amber-500/30 focus:ring-amber-500" :
                                   player.transferStatus === 'current' ? "bg-slate-900 border-white/5 text-slate-200 focus:ring-[#c5a85c]" :
                                   "bg-red-500/10 border-red-500/50 text-red-300 focus:ring-red-500"
                                 )}
                               >
                                 <option value="" disabled className="text-slate-500 bg-slate-950">ВЫБЕРИТЕ СТАТУС Изменения...</option>
                                 <option value="current" className="text-slate-200 bg-slate-950">Игрок текущего состава</option>
                                 <option value="new" className="text-cyan-400 bg-slate-950">Новый игрок (относительно отбора)</option>
                                 <option value="other_zone" className="text-amber-400 bg-slate-950">Заигран в другой зоне</option>
                               </select>
                             </div>
                           )}
                         </div>
                       </div>
                    </div>
                    );
                  })
                )}
             </div>

             {/* Кнопка "Добавить футболиста" снизу списка */}
             <div className="mt-5">
               <button 
                 type="button" 
                 onClick={handleAddPlayer}
                 className="w-full flex items-center justify-center gap-2 bg-gradient-to-r from-blue-900 to-blue-950 border border-blue-800 hover:from-blue-800 hover:to-blue-900 text-white hover:text-[#c5a85c] hover:border-[#c5a85c]/30 py-4.5 rounded-xl text-sm font-bold transition-all shadow-lg active:scale-95 cursor-pointer group"
               >
                 <UserPlus className="w-5 h-5 text-blue-400 group-hover:text-[#c5a85c] transition-colors" /> 
                 Добавить нового игрока в состав
               </button>
             </div>

             {form.players.some(p => p.status === 'previous') && (
               <div className="mt-6 flex items-start gap-3 bg-yellow-500/10 border border-yellow-500/20 p-4 rounded-xl text-yellow-100 text-xs leading-relaxed">
                 <AlertTriangle className="w-5 h-5 flex-shrink-0 text-yellow-400 mt-0.5" />
                 <div>
                   <p className="font-bold text-yellow-400 mb-1">Игроки прошлого состава загружены!</p>
                   <p className="opacity-80">
                     Футболисты, выделенные <span className="text-yellow-400 font-bold">жёлтой рамкой слева</span>, были автозагружены из базы прошлых сезонов. Вы можете свободно изменять их анкеты, убирать лишних игроков кнопкой (✖) или добавлять новых в состав через кнопку выше.
                   </p>
                 </div>
               </div>
             )}
          </div>

          <div className="pt-6 border-t border-white/5">
            <button
               type="submit"
               disabled={submitting}
               className="w-full relative overflow-hidden group bg-gradient-to-r from-[#c5a85c] to-[#e4cb8c] text-slate-950 font-extrabold flex items-center justify-center py-4 rounded-2xl shadow-[0_0_40px_rgba(197,168,92,0.25)] hover:scale-[1.01] active:scale-[0.99] transition-all disabled:opacity-75 disabled:cursor-not-allowed cursor-pointer text-base"
            >
              <div className="absolute inset-0 w-full h-full bg-gradient-to-r from-transparent via-white/20 to-transparent -translate-x-[100%] group-hover:translate-x-[100%] transition-transform duration-1000" />
              {submitting ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin mr-3 text-slate-950" />
                  Обработка и отправка заявки...
                </>
              ) : (
                <>
                  Подать официальную заявку на турнир
                  <Send className="w-5 h-5 ml-3" />
                </>
              )}
            </button>
            <div className="text-center text-xs text-slate-400 mt-5 leading-relaxed">
              * Заявка мгновенно запишется в централизованную Google Таблицу организационного комитета, а копия направится на email-адрес лиги.
              <br />
              <span className="text-slate-500 block mt-1">
                Для последующего ведения черновика просто возвращайтесь к этой веб-ссылке со своего Google-аккаунта.
              </span>
            </div>
          </div>
        </form>
        </>
        )}
      </main>

      <footer className="w-full max-w-4xl mx-auto text-center py-8 text-[11px] text-slate-500 border-t border-white/5 mt-12 pb-16">
        <div>© 2026 Ногайская Футбольная Лига (НФЛ). Все права защищены.</div>
        <div className="mt-2 flex justify-center gap-4">
          <a href="?admin=true" className="hover:text-[#c5a85c] text-indigo-400 hover:underline transition-all">
            Панель управления оргкомитета (Admin Panel)
          </a>
        </div>
      </footer>

      {/* Error / Validation Feedback Popup Modal */}
      {errorModal.isOpen && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
          {/* Overlay backdrop with smooth fade */}
          <div 
            className="absolute inset-0 bg-slate-950/80 backdrop-blur-sm transition-opacity"
            onClick={() => setErrorModal(prev => ({ ...prev, isOpen: false }))}
          />
          
          {/* Modal body card with bounce transition feel */}
          <div className="bg-slate-900 border border-red-500/20 rounded-2xl p-6 md:p-8 max-w-md w-full shadow-[0_20px_50px_rgba(239,68,68,0.15)] relative z-10 animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-start gap-4 mb-4">
              <div className="bg-red-500/10 text-red-400 p-2.5 rounded-xl border border-red-500/20">
                <AlertTriangle className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-lg font-bold text-white tracking-tight">{errorModal.title}</h3>
                <p className="text-xs text-slate-400 mt-1 font-mono uppercase tracking-wider text-[10px]">Валидация заявки нфл</p>
              </div>
            </div>

            <p className="text-sm text-slate-300 leading-relaxed mb-6">
              {errorModal.message}
            </p>

            {errorModal.steps && errorModal.steps.length > 0 && (
              <div className="bg-slate-950/55 rounded-xl p-4 border border-white/5 space-y-3 mb-6">
                <span className="text-[10px] text-slate-500 uppercase tracking-widest font-bold block mb-1">Как исправить ошибку:</span>
                {errorModal.steps.map((step, sIdx) => (
                  <div key={sIdx} className="flex items-start gap-2.5 text-xs text-slate-300 leading-relaxed">
                    <span className="text-red-400 font-mono font-bold mt-0.5">{sIdx + 1}.</span>
                    <span>{step}</span>
                  </div>
                ))}
              </div>
            )}

            <button
              type="button"
              onClick={() => setErrorModal(prev => ({ ...prev, isOpen: false }))}
              className="w-full bg-slate-800 hover:bg-slate-700 text-slate-200 font-bold py-3 px-4 rounded-xl transition-all text-xs cursor-pointer border border-white/5 text-center block"
            >
              Я понял, исправлю
            </button>
          </div>
        </div>
      )}

      {/* Fixed Floating Add Player Button */}
      {!showAdminPanel && (
        <div className="fixed bottom-6 right-6 z-[80] group md:bottom-8 md:right-8">
          <button
            type="button"
            onClick={handleAddPlayer}
            className="flex items-center justify-center w-14 h-14 rounded-full bg-gradient-to-r from-blue-600 to-blue-700 text-white shadow-[0_10px_25px_rgba(37,99,235,0.4)] hover:shadow-[0_12px_30px_rgba(37,99,235,0.6)] border border-blue-500/30 hover:scale-110 active:scale-95 transition-all duration-300 cursor-pointer"
            title="Быстро добавить игрока"
          >
            <Plus className="w-7 h-7" />
          </button>
          
          {/* Tooltip on hover */}
          <span className="absolute right-16 top-1/2 -translate-y-1/2 bg-slate-900 border border-white/10 text-slate-200 text-[11px] font-bold px-3 py-1.5 rounded-lg opacity-0 pointer-events-none group-hover:opacity-100 transition-opacity duration-200 whitespace-nowrap shadow-xl">
            Добавить игрока
          </span>
        </div>
      )}
    </div>

    {/* Elegant vector-printable layout for official roster PDF export */}
    {printSubmissions.length > 0 && (
      <div className="bg-white text-black min-h-screen p-8 font-sans" id="nfl-print-area">
        {printSubmissions.map((sub, sIdx) => (
          <div 
            key={sub.id || sIdx} 
            className="p-4 bg-white text-black max-w-[800px] mx-auto relative break-after-page"
            style={{ pageBreakBefore: sIdx > 0 ? 'always' : 'auto' }}
          >
            {/* Decorative federation-style header frame */}
            <div className="border-[3px] border-black p-5 mb-6 relative">
              <div className="absolute top-1 left-1 right-1 bottom-1 border border-black pointer-events-none" />
              <div className="text-center relative z-10">
                <h1 className="text-xl font-black uppercase tracking-wider text-black">
                  НОГАЙСКАЯ ФУТБОЛЬНАЯ ЛИГА
                </h1>
                <h2 className="text-xs font-bold tracking-widest text-slate-800 uppercase mt-0.5">
                  Официальный заявочный лист • Сезон 2026
                </h2>
                <div className="w-20 h-[1.5px] bg-black mx-auto my-2.5" />
                <p className="text-[9px] uppercase font-mono tracking-widest text-slate-500">
                  {(sub.stage || 'qualifier') === 'final' ? 'ФИНАЛЬНЫЙ ЭТАП' : 'ОТБОРОЧНЫЙ ЭТАП'} • ФЕДЕРАЦИЯ ЛЮБИТЕЛЬСКОГО ФУТБОЛА
                </p>
              </div>
            </div>

            {/* Team info segment */}
            <div className="grid grid-cols-2 gap-x-8 gap-y-2 mb-6 bg-slate-50 p-3 border border-black/10 rounded-sm text-xs">
              <div>
                <span className="text-slate-500 uppercase font-mono text-[8px] block">Футбольный клуб</span>
                <strong className="text-base text-black font-extrabold">{sub.teamName}</strong>
              </div>
              <div>
                <span className="text-slate-500 uppercase font-mono text-[8px] block">Район / Зона</span>
                <strong className="text-sm text-black font-bold">{sub.zone}</strong>
              </div>
              <div>
                <span className="text-slate-500 uppercase font-mono text-[8px] block">Капитан команды</span>
                <strong className="text-sm text-black font-bold">{sub.captainName}</strong>
              </div>
              <div>
                <span className="text-slate-500 uppercase font-mono text-[8px] block">Контактный телефон</span>
                <strong className="text-sm text-black font-mono font-bold">{sub.captainPhone}</strong>
              </div>
              <div>
                <span className="text-slate-500 uppercase font-mono text-[8px] block">Дата подачи</span>
                <span className="text-xs text-black font-medium">{formatRussianDate(sub.createdAt, true)}</span>
              </div>
              <div>
                <span className="text-slate-500 uppercase font-mono text-[8px] block">Состояние проверки</span>
                <strong className={`text-xs uppercase tracking-wider ${sub.synced ? "text-emerald-700" : "text-amber-600"}`}>
                  {sub.synced ? "СОГЛАСОВАНО" : "НА ПРОВЕРКЕ ОРГКОМИТЕТА"}
                </strong>
              </div>
            </div>

            {/* Players Table */}
            <div className="mb-8">
              <h3 className="text-[11px] font-black uppercase tracking-wider text-black mb-2 pb-1 border-b-[2px] border-black flex justify-between">
                <span>СОСТАВ КОМАНДЫ ({sub.players.length} ИГРОКОВ)</span>
                <span className="text-[9px] font-mono text-slate-500">Заявочный лист</span>
              </h3>
              <table className="w-full text-left border-collapse border border-black/20 text-[11px]">
                <thead>
                  <tr className="bg-slate-100 border-b border-black">
                    <th className="py-1.5 px-1 border-r border-black/20 text-center w-8">№</th>
                    <th className="py-1.5 px-2 border-r border-black/20">ФИО Игрока</th>
                    <th className="py-1.5 px-2 border-r border-black/20 w-24 text-center">Дата рождения</th>
                    <th className="py-1.5 px-2 border-r border-black/20 w-24 text-center">Амплуа</th>
                    <th className="py-1.5 px-2 border-r border-black/20 text-center w-12">Номер</th>
                    <th className="py-1.5 px-2 border-r border-black/20 text-center w-20">Тип</th>
                    {(sub.stage || 'qualifier') === 'final' && (
                      <th className="py-1.5 px-2 text-center w-24 border-black/20">Статус (Финал)</th>
                    )}
                  </tr>
                </thead>
                <tbody className="divide-y divide-black/10">
                  {sub.players.map((p, pIdx) => (
                    <tr key={p.id || pIdx} className="hover:bg-slate-50">
                      <td className="py-1.5 px-1 border-r border-black/10 text-center font-mono text-slate-500">{pIdx + 1}</td>
                      <td className="py-1.5 px-2 border-r border-black/10 font-bold text-black">{p.fullName}</td>
                      <td className="py-1.5 px-2 border-r border-black/10 text-center font-mono text-slate-700">{p.birthDate || '-'}</td>
                      <td className="py-1.5 px-2 border-r border-black/10 text-center text-slate-800">{p.position || '-'}</td>
                      <td className="py-1.5 px-2 border-r border-black/10 text-center font-mono font-bold text-black">#{p.number || '-'}</td>
                      <td className={`py-1.5 px-2 text-center text-[10px] font-mono ${(sub.stage || 'qualifier') === 'final' ? 'border-r border-black/10' : ''}`}>
                        {p.isLegionnaire ? (
                          <strong className="text-red-700">ЛЕГ</strong>
                        ) : (
                          <span className="text-slate-500">ОБЫЧНЫЙ</span>
                        )}
                        {p.isVerified ? (
                          <span className="text-emerald-700 font-extrabold ml-1" title="Заигран">✓</span>
                        ) : (
                          p.fullName && dbPlayers.some(db => db.fullName.toLowerCase() === p.fullName?.toLowerCase()) && (
                            <FileText className="w-3 h-3 text-orange-500 inline-block ml-1" title="Был в базе, но не заигран" />
                          )
                        )}
                      </td>
                      {(sub.stage || 'qualifier') === 'final' && (
                        <td className="py-1.5 px-2 text-center text-[10px] font-mono font-bold text-slate-700">
                          {p.transferStatus === 'new' ? 'НОВЫЙ' : p.transferStatus === 'other_zone' ? 'ДР. ЗОНА' : 'ТЕКУЩИЙ'}
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Captain and Commissioner Sign-off signatures */}
            <div className="mt-10 pt-4 border-t border-black/20 text-[11px]">
              <div className="grid grid-cols-2 gap-10">
                <div className="space-y-4">
                  <p className="text-slate-800 font-bold">Заявитель от клуба (Капитан):</p>
                  <div className="space-y-1">
                    <div className="w-full border-b border-black h-8 flex items-end justify-between px-1 pb-0.5">
                      <span className="text-[8px] text-slate-400 italic">Подпись</span>
                      <span className="text-black font-semibold text-xs">{sub.captainName}</span>
                    </div>
                    <div className="flex justify-between text-[8px] text-slate-500 font-mono">
                      <span>(подпись)</span>
                      <span>(расшифровка подписи)</span>
                    </div>
                  </div>
                  <p className="text-[9px] text-slate-400">Дата: "______" ____________________ 2026 г.</p>
                </div>
                <div className="space-y-4">
                  <p className="text-slate-800 font-bold">Оргкомитет лиги (НФЛ):</p>
                  <div className="space-y-1">
                    <div className="w-full border-b border-black h-8 flex items-end justify-between px-1 pb-0.5">
                      <span className="text-[8px] text-slate-400 italic">М.П. / Печать</span>
                      <span className="text-slate-400">___________________________</span>
                    </div>
                    <div className="flex justify-between text-[8px] text-slate-500 font-mono">
                      <span>(М.П.)</span>
                      <span>(расшифровка подписи представителя)</span>
                    </div>
                  </div>
                  <p className="text-[9px] text-slate-400">
                    Согласовано: {sub.synced ? formatRussianDate(sub.createdAt) : "______ ____________________ 2026 г."}
                  </p>
                </div>
              </div>
            </div>
          </div>
        ))}
      </div>
    )}
    </>
  );
}
