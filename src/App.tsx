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

  return (
    <ApplicationForm 
      isGuest={isGuest || (!user)}
      user={user}
      defaultShowAdmin={isAdminRequest && !!user}
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

