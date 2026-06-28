"use client";

import React, { useState } from "react";
import { User, Lock, LogIn, Eye, EyeOff, DollarSign, KeyRound, Loader2, AlertCircle } from "lucide-react";

export default function Login() {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError("");
    setLoading(true);

    try {
      const response = await fetch("/api/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username, password }),
      });

      const data = await response.json();

      if (!response.ok) {
        setError(data.error || "Credenciales incorrectas");
        setLoading(false);
      } else {
        window.location.reload();
      }
    } catch (err) {
      setError("Error de conexión al servidor");
      setLoading(false);
    }
  };

  return (
    <main className="min-h-screen w-full flex items-center justify-center bg-slate-50 p-4 relative select-none">
      <section className="w-full max-w-md bg-white border border-slate-200 rounded-3xl p-8 shadow-xl shadow-slate-100 relative z-10 overflow-hidden transition-all duration-300">
        <div className="mb-8 flex justify-center">
          <div className="text-center">
            <h1 className="text-2xl font-extrabold text-slate-900 tracking-tight leading-tight">
              Kimche Analyzer
            </h1>
            <p className="text-sm text-slate-500 mt-1">
              Convierte datos en información
            </p>
          </div>
        </div>

        <form onSubmit={handleSubmit} className="space-y-5">
          {error && (
            <div className="flex items-center gap-3 p-3.5 rounded-xl bg-red-50 border border-red-200 text-red-750 text-xs animate-shake">
              <AlertCircle className="h-4.5 w-4.5 shrink-0" />
              <span>{error}</span>
            </div>
          )}

          <div className="space-y-2">
            <label htmlFor="username-input" className="text-xs font-semibold text-slate-500 tracking-wide uppercase">
              Usuario
            </label>
            <div className="relative rounded-xl border border-slate-300 bg-slate-50 focus-within:bg-white focus-within:border-indigo-600 focus-within:ring-2 focus-within:ring-indigo-600/10 transition-all duration-300 shadow-sm">
              <User className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 h-4.5 w-4.5" />
              <input
                id="username-input"
                type="text"
                name="username"
                autoComplete="username"
                required
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="Ingresa tu usuario"
                className="w-full bg-transparent pl-11 pr-4 py-3 text-slate-800 placeholder-slate-450 focus:outline-none text-sm transition-all"
                disabled={loading}
              />
            </div>
          </div>

          <div className="space-y-2">
            <label htmlFor="current-password" className="text-xs font-semibold text-slate-500 tracking-wide uppercase">
              Contraseña
            </label>
            <div className="relative rounded-xl border border-slate-300 bg-slate-50 focus-within:bg-white focus-within:border-indigo-600 focus-within:ring-2 focus-within:ring-indigo-600/10 transition-all duration-300 shadow-sm">
              <Lock className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 h-4.5 w-4.5" />
              <input
                id="current-password"
                type={showPassword ? "text" : "password"}
                name="password"
                autoComplete="current-password"
                required
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="••••••••"
                className="w-full bg-transparent pl-11 pr-11 py-3 text-slate-800 placeholder-slate-450 focus:outline-none text-sm transition-all"
                disabled={loading}
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-650 transition-colors"
                tabIndex={-1}
              >
                {showPassword ? <EyeOff className="h-4.5 w-4.5" /> : <Eye className="h-4.5 w-4.5" />}
              </button>
            </div>
          </div>

          <div className="space-y-3 pt-2">
            <button
              type="submit"
              disabled={loading}
              className="w-full flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-3 px-4 rounded-xl shadow-md shadow-indigo-600/10 transition-all duration-200 transform active:scale-[0.98] cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed text-sm"
            >
              {loading ? (
                <Loader2 className="h-4.5 w-4.5 animate-spin" />
              ) : (
                <>
                  <LogIn className="h-4.5 w-4.5" />
                  <span>Iniciar sesión</span>
                </>
              )}
            </button>

            <a
              href="https://wa.me/56934417698?text=%C2%A1Hola%21%20Me%20gustar%C3%ADa%20conseguir%20credenciales%20para%20Kimche%20Analyzer.%20%C2%BFC%C3%B3mo%20lo%20hago%3F"
              target="_blank"
              rel="noopener noreferrer"
              className="w-full inline-flex items-center justify-center gap-2 px-4 py-3 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold tracking-wide transition-all duration-300 shadow-md shadow-emerald-600/10 cursor-pointer"
            >
              <DollarSign className="h-4.5 w-4.5" />
              Consigue acceso por $2.000 (pago único)
            </a>
          </div>
        </form>
      </section>
    </main>
  );
}
