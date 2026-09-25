import React, { useState } from 'react';
import { LockKeyhole } from 'lucide-react';
import { supabase } from './supabaseClient';

export default function Login() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setIsLoading(true);
    const { error: signInError } = await supabase.auth.signInWithPassword({ email, password });
    setIsLoading(false);
    if (signInError) {
      setError('Correo o contraseña incorrectos.');
    }
  };

  return (
    <div className="flex h-screen items-center justify-center bg-slate-50 dark:bg-slate-950 px-4">
      <form
        onSubmit={handleSubmit}
        className="w-full max-w-sm bg-white dark:bg-slate-800 p-8 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700"
      >
        <div className="flex flex-col items-center mb-6">
          <div className="w-12 h-12 rounded-xl bg-emerald-100 dark:bg-emerald-900/30 flex items-center justify-center mb-3">
            <LockKeyhole className="text-emerald-600" size={22} />
          </div>
          <h1 className="text-lg font-bold text-slate-800 dark:text-white text-center">Grupo Procreas 929, C.A.</h1>
          <p className="text-xs text-slate-500 text-center mt-1">Contabilidad NIIF — Acceso privado</p>
        </div>

        {error && (
          <div className="mb-4 text-xs bg-red-50 dark:bg-red-900/20 text-red-700 dark:text-red-300 border border-red-200 dark:border-red-800 rounded-lg px-3 py-2">
            {error}
          </div>
        )}

        <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-1">Correo</label>
        <input
          type="email"
          required
          value={email}
          onChange={e => setEmail(e.target.value)}
          className="w-full mb-4 p-2.5 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none"
          placeholder="tucorreo@ejemplo.com"
        />

        <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-1">Contraseña</label>
        <input
          type="password"
          required
          value={password}
          onChange={e => setPassword(e.target.value)}
          className="w-full mb-6 p-2.5 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none"
          placeholder="••••••••"
        />

        <button
          type="submit"
          disabled={isLoading}
          className="w-full bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white text-sm font-medium py-2.5 rounded-lg transition-colors"
        >
          {isLoading ? 'Entrando...' : 'Entrar'}
        </button>

        <p className="text-[11px] text-slate-400 text-center mt-4">
          Acceso restringido. Si no tienes credenciales, pídeselas al administrador del sistema.
        </p>
      </form>
    </div>
  );
}
