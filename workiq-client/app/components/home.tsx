import type { AccountInfo } from "@azure/msal-browser";
import { WorkIqChatPanel } from "./workiq-chat";

export function HomeShell({ children }: { children: React.ReactNode }) {
  return (
    <main className="relative flex min-h-screen items-center justify-center overflow-hidden bg-slate-950 px-4 py-12 text-white sm:px-6">
      <div
        aria-hidden="true"
        className="pointer-events-none absolute inset-0 overflow-hidden"
      >
        <div className="absolute -left-32 -top-40 h-96 w-96 rounded-full bg-indigo-500/20 blur-3xl" />
        <div className="absolute -bottom-48 -right-24 h-[32rem] w-[32rem] rounded-full bg-cyan-400/10 blur-3xl" />
      </div>
      <section className="relative w-full max-w-xl rounded-3xl border border-white/10 bg-slate-900/80 p-6 shadow-2xl shadow-black/30 backdrop-blur-xl sm:p-10">
        {children}
      </section>
    </main>
  );
}

export function HomeHeader() {
  return (
    <>
      <div className="mb-9 flex items-center gap-3">
        <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-indigo-400/15 text-indigo-300 ring-1 ring-inset ring-indigo-300/20">
          <svg
            aria-hidden="true"
            className="h-6 w-6"
            fill="none"
            viewBox="0 0 24 24"
          >
            <path
              d="M4 5.5h7v5H4v-5Zm9 0h7v8h-7v-8ZM4 12.5h7v6H4v-6Zm9 3h7v3h-7v-3Z"
              fill="currentColor"
            />
          </svg>
        </div>
        <span className="text-sm font-semibold tracking-wide text-slate-200">
          Work IQ
        </span>
      </div>
      <h1 className="text-3xl font-semibold tracking-tight text-white sm:text-4xl">
        Welcome to Work IQ
      </h1>
    </>
  );
}

export function AuthProgress() {
  return (
    <div
      aria-live="polite"
      className="mt-8 flex items-center gap-3 rounded-2xl border border-indigo-300/15 bg-indigo-300/5 p-4 text-sm text-indigo-100"
      role="status"
    >
      <span className="h-5 w-5 animate-spin rounded-full border-2 border-indigo-200/30 border-t-indigo-200" />
      Authentication in progress…
    </div>
  );
}

export function SignInPanel({ onSignIn }: { onSignIn: () => void }) {
  return (
    <div className="mt-8">
      <div className="rounded-2xl border border-white/10 bg-white/[0.03] p-4">
        <p className="text-sm leading-6 text-slate-300">
          You’re not signed in yet. Use your organization’s Microsoft account
          to securely access Work IQ.
        </p>
      </div>
      <button
        className="mt-5 inline-flex min-h-12 w-full items-center justify-center gap-3 rounded-xl bg-indigo-400 px-5 py-3 text-sm font-semibold text-slate-950 shadow-lg shadow-indigo-950/30 transition hover:bg-indigo-300 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-indigo-200"
        onClick={onSignIn}
        type="button"
      >
        <MicrosoftMark />
        Sign in with Microsoft
      </button>
    </div>
  );
}

export function SignedInPanel({
  account,
  onSignOut,
}: {
  account: AccountInfo;
  onSignOut: () => void;
}) {
  return (
    <div className="mt-8">
      <div className="mb-5 inline-flex items-center gap-2 rounded-full bg-emerald-400/10 px-3 py-1.5 text-sm font-medium text-emerald-300 ring-1 ring-inset ring-emerald-300/20">
        <span className="h-1.5 w-1.5 rounded-full bg-emerald-300" />
        Signed in
      </div>
      <dl className="divide-y divide-white/10 rounded-2xl border border-white/10 bg-white/[0.03]">
        <AccountDetail label="Name" value={account.name ?? "—"} />
        <AccountDetail label="Username" value={account.username} />
        <AccountDetail label="Tenant ID" value={account.tenantId} monospace />
      </dl>
      <WorkIqChatPanel account={account} key={account.homeAccountId} />
      <button
        className="mt-6 inline-flex min-h-12 w-full items-center justify-center rounded-xl border border-white/15 bg-white/5 px-5 py-3 text-sm font-semibold text-white transition hover:border-white/25 hover:bg-white/10 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-indigo-300"
        onClick={onSignOut}
        type="button"
      >
        Sign out
      </button>
    </div>
  );
}

export function HomeFooter() {
  return (
    <p className="mt-8 border-t border-white/10 pt-5 text-xs leading-5 text-slate-500">
      Secure sign-in powered by Microsoft.
    </p>
  );
}

function AccountDetail({
  label,
  value,
  monospace = false,
}: {
  label: string;
  value: string;
  monospace?: boolean;
}) {
  return (
    <div className="grid gap-1 px-4 py-3 sm:grid-cols-[8rem_1fr] sm:gap-4">
      <dt className="text-sm text-slate-400">{label}</dt>
      <dd
        className={`break-all text-sm text-slate-100 ${
          monospace ? "font-mono text-xs text-slate-300" : "font-medium"
        }`}
      >
        {value}
      </dd>
    </div>
  );
}

function MicrosoftMark() {
  return (
    <svg
      aria-hidden="true"
      className="h-4 w-4"
      fill="currentColor"
      viewBox="0 0 16 16"
    >
      <path d="M1 1h6.5v6.5H1V1Zm7.5 0H15v6.5H8.5V1ZM1 8.5h6.5V15H1V8.5Zm7.5 0H15V15H8.5V8.5Z" />
    </svg>
  );
}
