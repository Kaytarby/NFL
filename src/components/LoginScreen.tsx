import React from 'react';
import { googleSignIn } from '../lib/firebase/firebase';
import { Sparkles, Trophy } from 'lucide-react';

interface LoginScreenProps {
  onLogin: () => void;
  onGuestAccess: () => void;
}

export default function LoginScreen({ onLogin, onGuestAccess }: LoginScreenProps) {
  const handleLogin = async () => {
    try {
      const result = await googleSignIn();
      if (result) {
        onLogin();
      }
    } catch (err) {
      console.error('Login failed:', err);
      alert('Ошибка при авторизации. Пожалуйста, убедитесь, что вы подтвердили необходимые разрешения для работы с Google Таблицами.');
    }
  };

  return (
    <div className="min-h-screen bg-[#020516] star-bg text-white flex flex-col items-center justify-center p-4 relative overflow-hidden">
      
      {/* UEFA Champions League ambient blue glowing circles */}
      <div className="absolute top-1/4 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[500px] h-[500px] bg-gradient-to-tr from-blue-700/20 to-cyan-500/10 rounded-full filter blur-[120px] pointer-events-none" />
      <div className="absolute bottom-10 right-10 w-[300px] h-[300px] bg-gradient-to-tr from-[#c5a85c]/10 to-transparent rounded-full filter blur-[100px] pointer-events-none" />

      <div className="max-w-md w-full bg-slate-950/70 backdrop-blur-xl p-8 rounded-2xl border border-slate-800 shadow-[0_0_50px_rgba(0,0,0,0.8)] relative z-10">
        
        {/* Glow accent ribbon */}
        <div className="absolute top-0 left-0 right-0 h-[2px] bg-gradient-to-r from-blue-500 via-[#c5a85c] to-cyan-500" />
        
        <div className="mb-6 flex justify-center">
          <div className="bg-slate-900 border border-slate-800 p-5 rounded-2xl shadow-[0_0_20px_rgba(197,168,92,0.15)]">
            <Trophy className="w-16 h-16 text-[#c5a85c]" />
          </div>
        </div>
        
        <div className="text-center mb-8">
          <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-blue-500/10 text-cyan-400 text-xs font-semibold rounded-full uppercase tracking-wider mb-3 border border-blue-500/20">
            <Sparkles className="w-3.5 h-3.5 text-[#c5a85c]" /> ОТБОР В ЛИГУ
          </span>
          <h1 className="text-2xl font-bold tracking-tight text-white font-display">
            Ногайская Футбольная Лига
          </h1>
          <p className="text-sm text-slate-400 mt-2">
            Заявка на участие в квалификационном отборе
          </p>
        </div>
        
        <div className="space-y-3">
          <button 
            onClick={onGuestAccess}
            className="w-full flex items-center justify-center gap-3 bg-gradient-to-r from-[#c5a85c] to-[#e4cb8c] text-slate-950 px-6 py-4 rounded-xl font-bold text-base shadow-[0_0_20px_rgba(197,168,92,0.3)] hover:scale-[1.02] active:scale-[0.98] transition-all cursor-pointer focus:ring-2 focus:ring-[#c5a85c] focus:outline-none"
          >
            Заполнить в режиме Гостя
          </button>

          <button 
            type="button"
            onClick={handleLogin}
            className="w-full flex items-center justify-center gap-3 bg-slate-900 border border-slate-850 hover:border-slate-700 text-slate-300 hover:text-white px-6 py-3.5 rounded-xl font-medium text-sm hover:scale-[1.01] active:scale-[0.99] transition-all cursor-pointer focus:ring-2 focus:ring-slate-700 focus:outline-none"
          >
            <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" alt="Google" className="w-4 h-4 opacity-80" />
            Войти по Google (Организаторы)
          </button>
        </div>

        <p className="text-xs text-slate-400 text-center mt-6 leading-relaxed">
          Вы можете подать заявку в Лигу напрямую без авторизации. Авторизация по Google рекомендуется организаторам и пользователям, желающим управлять прошлыми заявками.
        </p>

      </div>
    </div>
  );
}

