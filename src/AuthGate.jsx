import React, { useEffect, useState } from 'react';
import { supabase } from './supabaseClient';
import Login from './Login.jsx';

// Envuelve la app: mientras no haya una sesión válida de Supabase Auth,
// muestra la pantalla de inicio de sesión en vez del contenido.
export default function AuthGate({ children }) {
  const [session, setSession] = useState(undefined); // undefined = todavía no se sabe

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));

    const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession);
    });

    return () => listener.subscription.unsubscribe();
  }, []);

  if (session === undefined) {
    return (
      <div className="flex h-screen items-center justify-center bg-slate-50 dark:bg-slate-950">
        <p className="text-slate-400 text-sm">Verificando sesión...</p>
      </div>
    );
  }

  if (!session) {
    return <Login />;
  }

  return children;
}
