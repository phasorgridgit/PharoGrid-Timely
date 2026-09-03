import { useEffect, useState } from "react";
import TopNav from "./components/TopNav";
import Logo from "./components/Logo";
import Dashboard from "./pages/Dashboard";
import Calendar from "./pages/Calendar";
import Timesheets from "./pages/Timesheets";
import Reports from "./pages/Reports";
import TimeOff from "./pages/TimeOff";
import ClientsProjects from "./pages/ClientsProjects";
import Team from "./pages/Team";
import Billing from "./pages/Billing";
import UpdateBanner from "./components/UpdateBanner";
import { isMockMode } from "./graph/graphClient";
import { getCurrentEmployee, resetCurrentEmployeeCache } from "./graph/currentUser";

const PAGES = {
  dashboard: Dashboard,
  calendar: Calendar,
  timesheets: Timesheets,
  reports: Reports,
  timeoff: TimeOff,
  clients: ClientsProjects,
  team: Team,
  billing: Billing
};

export default function App() {
  const [account, setAccount] = useState(null);
  const [employeeName, setEmployeeName] = useState(null);
  const [employeeError, setEmployeeError] = useState(null);
  const [page, setPage] = useState("dashboard");

  useEffect(() => {
    if (isMockMode) {
      // No Microsoft sign-in, no Electron IPC call — just drop straight into
      // the app with a fake local account. Nothing is saved anywhere.
      setAccount({ signedIn: true, username: "demo@phasorgrid.local (mock mode)" });
      return;
    }
    window.workhub.getToken().then((token) => {
      // A cached token on launch means silent SSO worked — no prompt needed.
      if (token) setAccount({ signedIn: true });
    });
  }, []);

  useEffect(() => {
    if (!account) return;
    getCurrentEmployee()
      .then((e) => setEmployeeName(e.name))
      .catch((err) => setEmployeeError(err.message));
  }, [account]);

  async function handleSignIn() {
    if (isMockMode) {
      resetCurrentEmployeeCache();
      setEmployeeError(null);
      setPage("dashboard");
      setAccount({ signedIn: true, username: "demo@phasorgrid.local (mock mode)" });
      return;
    }
    const result = await window.workhub.login();
    setAccount({ signedIn: true, username: result.account });
  }

  /**
   * The Sign out button used to call window.workhub.logout() directly with
   * no follow-up — that clears the MSAL token cache in the main process
   * (real fix, that part worked), but the renderer never learned the logout
   * happened: `account` stayed set, so the UI just sat on whatever page it
   * was already showing. In mock mode it was worse — window.workhub doesn't
   * exist there at all, so the click did nothing whatsoever.
   *
   * This also has to clear the cached employee resolution
   * (resetCurrentEmployeeCache) — getCurrentEmployee() caches "who am I"
   * for the life of the renderer, which is exactly what makes concurrent
   * multi-employee use safe, but it means a second person signing in after
   * a sign-out on the same running window would otherwise still resolve to
   * the first person's identity.
   */
  async function handleSignOut() {
    if (!isMockMode) {
      // If clearing the MSAL token cache in the Electron main process
      // fails for any reason (a stale/mismatched cached account, a
      // transient IPC error, etc.), `window.workhub.logout()` rejects —
      // and an unhandled rejection here used to abort this whole function
      // before any of the state resets below ran. From the person's side
      // that looked exactly like "the Sign out button does nothing":
      // the click registered, but the screen never changed. A failure to
      // clear the *background* token cache should never block the
      // person from actually being signed out of the app they're looking
      // at, so this is now best-effort: log it, but always fall through
      // to resetting local state. (Worst case if this particular clear
      // fails: silent SSO could re-authenticate them on next launch —
      // annoying, but nowhere near as bad as a Sign out button that
      // can't be clicked.)
      try {
        await window.workhub?.logout();
      } catch (err) {
        console.error("workhub.logout() failed — signing out of the app locally anyway:", err);
      }
    }
    resetCurrentEmployeeCache();
    setAccount(null);
    setEmployeeName(null);
    setEmployeeError(null);
    setPage("dashboard");
  }

  if (!account) {
    return (
      <div className="signin-screen">
        <Logo size={44} />
        <h1>PhasorGrid Timely</h1>
        <p>Sign in with your PhasorGrid Microsoft 365 account to continue.</p>
        <button onClick={handleSignIn}>Sign in with Microsoft</button>
      </div>
    );
  }

  if (employeeError) {
    return (
      <div className="signin-screen">
        <Logo size={44} />
        <h1>Account not set up yet</h1>
        <p>{employeeError}</p>
        <button
          onClick={() => {
            resetCurrentEmployeeCache();
            setEmployeeError(null);
            getCurrentEmployee().then((e) => setEmployeeName(e.name)).catch((err) => setEmployeeError(err.message));
          }}
        >
          Try again
        </button>
      </div>
    );
  }

  const Page = PAGES[page];

  return (
    <div className="shell">
      <TopNav current={page} onNavigate={setPage} username={account.username} employeeName={employeeName} onSignOut={handleSignOut} />
      <UpdateBanner />
      <main className="main">
        {isMockMode && (
          <div className="mock-banner">
            Mock mode — nothing here is saved. No Microsoft sign-in, no SharePoint. Data resets when you close the app.
          </div>
        )}
        <Page />
      </main>
    </div>
  );
}
