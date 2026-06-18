import { getAccessToken } from './firebase/firebase';

const SPREADSHEET_ID = (import.meta as any).env.VITE_SPREADSHEET_ID;

export const formatRussianDate = (dateString: string | Date, includeTime = false) => {
  if (!dateString) return '';
  const d = new Date(dateString);
  if (isNaN(d.getTime())) return String(dateString);
  const pad = (n: number) => String(n).padStart(2, '0');
  const dateFormatted = `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}`;
  if (includeTime) {
    return `${dateFormatted} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
  }
  return dateFormatted;
};

export interface Player {
  id: string; // generated
  teamName: string;
  fullName: string;
  birthDate: string;
  position: string;
  number: string;
  isConfirmed?: boolean;
  status: 'previous' | 'new' | 'deleted';
  isLegionnaire?: boolean;
  isVerified?: boolean; // "заигран" (documents verified)
  transferStatus?: 'current' | 'new' | 'other_zone';
}

export interface ApplicationDraft {
  teamName: string;
  zone: string;
  captainName: string;
  captainPhone: string;
  logoUrl: string | null;
  players: Player[];
  version?: number;
  stage?: 'qualifier' | 'final';
}

// Helper to parse sheet ranges like "'Teams'!A:A" or "Players!A:E"
const parseRange = (rangeStr: string): { sheet: string; range: string } => {
  const match = rangeStr.match(/['"]?([^'"]+)['"]?!(.+)/);
  if (match) {
    return { sheet: match[1], range: match[2] };
  }
  return { sheet: rangeStr, range: '' };
};

export const fetchPublicSheetData = async (rangeStr: string): Promise<string[][]> => {
  if (!SPREADSHEET_ID) throw new Error('VITE_SPREADSHEET_ID is not configured.');
  
  const { sheet, range } = parseRange(rangeStr);
  const rangeQuery = range ? `&range=${range}` : '';
  const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheet)}${rangeQuery}`;
  
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Failed to fetch public sheet data for ${sheet}`);
  }
  
  const text = await res.text();
  const jsonStart = text.indexOf('{');
  const jsonEnd = text.lastIndexOf('}');
  if (jsonStart === -1 || jsonEnd === -1) {
    throw new Error(`Invalid JSON format from google visualization API`);
  }
  
  const jsonStr = text.substring(jsonStart, jsonEnd + 1);
  const data = JSON.parse(jsonStr);
  
  if (data.status === 'error') {
    const errorDetails = data.errors?.map((err: any) => err.detailed_message || err.message).join(', ');
    throw new Error(`Ошибка Google Таблиц: ${errorDetails || 'Доступ ограничен (Access Denied) или лист не найден.'}`);
  }
  
  if (!data?.table?.rows) {
     return [];
  }

  const rows: string[][] = data.table.rows.map((r: any) => {
    if (!r || !r.c) return [];
    return r.c.map((cell: any) => {
      if (!cell) return '';
      if (cell.f !== undefined && cell.f !== null) return String(cell.f);
      if (cell.v !== null && cell.v !== undefined) {
        const v = String(cell.v);
        // Google Visualization API returns dates as 'Date(Year, Month, Day)' where Month is 0-indexed.
        const dateMatch = v.match(/^Date\((\d+),\s*(\d+),\s*(\d+)\)$/);
        if (dateMatch) {
          const year = dateMatch[1];
          const month = String(parseInt(dateMatch[2], 10) + 1).padStart(2, '0');
          const day = dateMatch[3].padStart(2, '0');
          return `${day}.${month}.${year}`;
        }
        return v;
      }
      return '';
    });
  });
  
  return rows;
};

export const fetchSheetData = async (range: string) => {
  if (!SPREADSHEET_ID) throw new Error('VITE_SPREADSHEET_ID is not configured.');
  
  let token = null;
  try {
    token = await getAccessToken();
  } catch (err) {
    console.log("Not logged in, proceeding with guest mode read fallback");
  }

  if (!token) {
    // FALLBACK: Fetch publicly without credentials (expects Google Sheet to be shared as "Anyone with link can view")
    console.log(`No active token. Attempting anonymous public read for ${range}`);
    return fetchPublicSheetData(range);
  }

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}?majorDimension=ROWS`;
  try {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });
    
    if (!res.ok) {
      console.warn(`Authenticated fetch failed with status ${res.status}, trying public fallback for range ${range}`);
      return fetchPublicSheetData(range);
    }

    const data = await res.json();
    return data.values || [];
  } catch (err) {
    console.warn("Error during authenticated fetch, falling back to public visualization query:", err);
    return fetchPublicSheetData(range);
  }
};

export const fetchTeams = async (): Promise<string[]> => {
  try {
    const rows = await fetchSheetData("'Teams'!A:A");
    return rows.slice(1).map(r => typeof r[0] === 'string' ? r[0].trim() : String(r[0] || '')).filter(Boolean);
  } catch (e: any) {
    console.warn("Failed to fetch teams, maybe sheet doesn't exist?", e);
    throw new Error(`Не удалось загрузить список команд. Убедитесь, что лист 'Teams' существует и доступ по ссылке открыт. Детали: ${e.message || e}`);
  }
};

export const fetchZones = async (): Promise<string[]> => {
  try {
    const rows = await fetchSheetData("'Zones'!A:A");
    return rows.slice(1).map(r => typeof r[0] === 'string' ? r[0].trim() : String(r[0] || '')).filter(Boolean);
  } catch (e: any) {
    console.warn("Failed to fetch zones, maybe sheet doesn't exist?", e);
    throw new Error(`Не удалось загрузить список районов. Убедитесь, что лист 'Zones' существует. Детали: ${e.message || e}`);
  }
};

export const fetchPlayers = async (): Promise<Omit<Player, 'id' | 'isConfirmed' | 'status'>[]> => {
  try {
    const rows = await fetchSheetData("'Players'!A:G");
    // Column order: Team Name, Player Name, Birth Date, Position, Number, isLegionnaire, isVerified
    return rows.slice(1).map(r => {
      const legRaw = r[5] != null ? String(r[5]).trim().toLowerCase() : '';
      const verRaw = r[6] != null ? String(r[6]).trim().toLowerCase() : '';
      return {
        teamName: r[0] != null ? String(r[0]).trim() : '',
        fullName: r[1] != null ? String(r[1]).trim() : '',
        birthDate: r[2] != null ? String(r[2]).trim() : '',
        position: r[3] != null ? String(r[3]).trim() : '',
        number: r[4] != null ? String(r[4]).trim() : '',
        isLegionnaire: legRaw === 'да' || legRaw === 'yes' || legRaw === '1' || legRaw === 'true',
        isVerified: verRaw === 'да' || verRaw === 'yes' || verRaw === '1' || verRaw === 'true'
      };
    }).filter(p => !!p.fullName);
  } catch (e: any) {
    console.warn("Failed to fetch players, maybe sheet doesn't exist?", e);
    throw new Error(`Не удалось загрузить игроков из базы. Убедитесь, что лист 'Players' существует. Детали: ${e.message || e}`);
  }
};

const ensureSheetExists = async (token: string, sheetName: string = 'Submissions') => {
  if (!SPREADSHEET_ID) return;
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}?fields=sheets.properties.title`;
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) {
      console.warn("Failed checking spreadsheet metadata, status:", res.status);
      return;
    }
    const meta = await res.json();
    const exists = meta.sheets?.some((s: any) => s.properties?.title === sheetName);
    
    if (exists) {
      return;
    }

    console.log(`Sheet '${sheetName}' not found, creating sheet '${sheetName}' dynamically...`);
    const updateUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}:batchUpdate`;
    const resCreate = await fetch(updateUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        requests: [
          {
            addSheet: {
              properties: {
                title: sheetName
              }
            }
          }
        ]
      })
    });

    if (!resCreate.ok) {
      console.warn(`Could not add sheet '${sheetName}':`, await resCreate.text());
      return; // If it fails, maybe it already exists but we couldn't see it (permissions?)
    }

    // Write localized headers
    const writeUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(sheetName)}!A1%3AM1?valueInputOption=USER_ENTERED`;
    const headers = [
      "Дата подачи", 
      "Название команды", 
      "Район / Зона", 
      "ФИО капитана", 
      "Телефон капитана", 
      "Ссылка на логотип", 
      "ФИО игрока", 
      "Дата рождения", 
      "Амплуа", 
      "Номер", 
      "Статус",
      "Легионер",
      "Заигран",
      "Версия"
    ];
    await fetch(writeUrl, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        values: [headers]
      })
    });
    console.log(`Dynamically created '${sheetName}' sheet with custom headers.`);
  } catch (err) {
    console.warn(`Error in ensureSheetExists for ${sheetName}, continuing to try the append:`, err);
  }
};

export const submitApplication = async (data: ApplicationDraft, sheetName: string = 'Submissions') => {
  if (!SPREADSHEET_ID) throw new Error('VITE_SPREADSHEET_ID is not configured.');
  const token = await getAccessToken();
  if (!token) throw new Error('Not authenticated');

  // Ensure table exists
  await ensureSheetExists(token, sheetName);

  // We write to a Submissions sheet. We can do an append operation.
  // Each row will be a player in the application.
  
  const values = data.players
    .filter(p => p.status !== 'deleted')
    .map(p => [
      formatRussianDate(new Date(), true),
      data.teamName,
      data.zone,
      data.captainName,
      data.captainPhone,
      data.logoUrl || '',
      p.fullName,
      p.birthDate,
      p.position,
      p.number,
      p.status,
      p.isLegionnaire ? 'Да' : 'Нет',
      p.isVerified ? 'Да' : 'Нет',
      data.version || 1
    ]);
    
  if (values.length === 0) {
     // Just append the team without players
     values.push([
       formatRussianDate(new Date(), true),
       data.teamName,
       data.zone,
       data.captainName,
       data.captainPhone,
       data.logoUrl || '',
       '', '', '', '', '', '', '',
       data.version || 1
     ]);
  }

  const range = encodeURIComponent(`${sheetName}!A:N`);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}:append?valueInputOption=USER_ENTERED`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      values
    })
  });

  if (!res.ok) {
    const err = await res.json();
    throw new Error(`Failed to submit application: ${err.error?.message || 'Unknown error'}`);
  }
  
  return true;
};
