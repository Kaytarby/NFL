/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { User } from 'firebase/auth';
import { initAuth, logout, auth } from './lib/firebase/firebase';
import LoginScreen from './components/LoginScreen';
import ApplicationForm from './components/ApplicationForm';
import { Loader2 } from 'lucide-react';

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [isGuest, setIsGuest] = useState(false);

  // List of emails allowed to enter the Admin Panel
  // We use lowercase to ensure match
  const ADMIN_EMAILS = (import.meta as any).env.VITE_ADMIN_EMAILS
    ? (import.meta as any).env.VITE_ADMIN_EMAILS.split(',').map((e: string) => e.trim().toLowerCase())
    : ['kaytarby.88@gmail.com', 'mbairam2107@gmail.com'];

  const isAdminRequest = typeof window !== 'undefined' && 
    (window.location.search.includes('admin=true') || window.location.search.includes('admin'));

  useEffect(() => {
    if (!isAdminRequest) {
      setIsGuest(true);
    }
    const unsubscribe = initAuth(
      (currentUser) => {
        setUser(currentUser);
        setLoading(false);
      },
      () => {
        setUser(null);
        setLoading(false);
      }
    );
    return () => {
      // @ts-ignore
      if (typeof unsubscribe === 'function') unsubscribe();
    };
  }, [isAdminRequest]);

  if (loading) {
    return (
      <div className="min-h-screen bg-[#020516] flex items-center justify-center">
        <Loader2 className="w-12 h-12 text-[#c5a85c] animate-spin" />
      </div>
    );
  }

  // Admin flow: requires signing in via Google, UNLESS they explicitly chose guest
  if (isAdminRequest && !user && !isGuest) {
    return (
      <LoginScreen 
        onLogin={() => setUser(auth.currentUser)} 
        onGuestAccess={() => {
           setIsGuest(true);
           // Clear admin search param so it doesn't linger
           window.history.replaceState({}, '', '/');
        }} 
      />
    );
  }

  // Double check if the logged in user is actually an admin
  const isActuallyAdmin = isAdminRequest && user && user.email && ADMIN_EMAILS.includes(user.email.toLowerCase());

  if (isAdminRequest && user && !isActuallyAdmin) {
    return (
      <div className="min-h-screen bg-[#020516] flex flex-col items-center justify-center p-4">
        <div className="bg-slate-900 border border-red-500/30 p-8 rounded-2xl max-w-md w-full text-center">
          <h2 className="text-xl font-bold text-white mb-2">Доступ запрещен</h2>
          <p className="text-slate-400 mb-6">Ваша почта ({user.email}) не зарегистрирована как администратор.</p>
          <button 
             onClick={() => {
                logout();
                window.history.replaceState({}, '', '/');
                window.location.reload();
             }}
             className="w-full bg-[#c5a85c] text-white font-medium py-3 rounded-xl hover:opacity-90"
          >
            Выйти и вернуться на главную
          </button>
        </div>
      </div>
    );
  }

  return (
    <ApplicationForm 
      isGuest={isGuest || (!user)}
      user={user}
      defaultShowAdmin={!!isActuallyAdmin}
      onLogout={() => {
        if (user) {
          logout();
        } else {
          setIsGuest(true);
        }
      }} 
    />
  );
}

