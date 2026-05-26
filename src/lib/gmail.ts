import { getAccessToken } from './firebase/firebase';

const ADMIN_EMAIL = (import.meta as any).env.VITE_ADMIN_EMAIL;

export const sendNotificationEmail = async (teamName: string, captainName: string, captainPhone: string) => {
  const token = await getAccessToken();
  if (!token) return;

  const toEmail = ADMIN_EMAIL;
  if (!toEmail) return; // If not configured, just skip silently or throw. Let's skip silently to not break submission.

  const subject = `Новая заявка: ${teamName}`;
  const text = `Поступила новая заявка на турнир от команды "${teamName}".
  
Капитан: ${captainName}
Телефон: ${captainPhone}

Детали заявки сохранены в Google Таблице.`;

  const emailLines = [
    `To: ${toEmail}`,
    `Subject: =?utf-8?B?${btoa(unescape(encodeURIComponent(subject)))}?=`,
    'Content-Type: text/plain; charset=utf-8',
    '',
    text
  ];
  
  const emailStr = emailLines.join('\r\n');
  const base64EncodedEmail = btoa(unescape(encodeURIComponent(emailStr)))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');

  const response = await fetch('https://gmail.googleapis.com/upload/gmail/v1/users/me/messages/send', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      raw: base64EncodedEmail,
    }),
  });

  if (!response.ok) {
    console.error('Failed to send email API status:', response.status);
  }
};
