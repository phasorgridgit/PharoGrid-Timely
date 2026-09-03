import brandLockup from "../assets/logo.png";
import { isMockMode } from "../graph/graphClient";

const NAV_ITEMS = [
  { key: "dashboard", label: "Dashboard" },
  { key: "calendar", label: "Calendar" },
  { key: "timesheets", label: "Timesheets" },
  { key: "reports", label: "Reports" },
  { key: "timeoff", label: "Time Off" },
  { key: "clients", label: "Clients & Projects" },
  { key: "team", label: "Team" },
  { key: "billing", label: "Billing & Invoices" }
];

export default function TopNav({ current, onNavigate, username, employeeName, onSignOut }) {
  const displayName = employeeName || username || "U";
  return (
    <div className="topnav-shell">
      <div className="topnav-brand-bar">
        <div className="topnav-appname">
          Timely
          <span>PhasorGrid</span>
        </div>

        <img className="topnav-logo" src={brandLockup} alt="PhasorGrid" />

        <div className="topnav-right">
          <div className="sync-chip">
            <span className={isMockMode ? "dot amber" : "dot"} />
            {isMockMode ? "Mock mode" : "Synced"}
          </div>
          <div className="user-chip">
            <div className="avatar">{displayName.slice(0, 2).toUpperCase()}</div>
            <div className="role" title={employeeName ? `Signed in as ${employeeName}` : undefined}>{displayName}</div>
            <div className="role" onClick={onSignOut} style={{ cursor: "pointer", marginLeft: 8 }}>
              Sign out
            </div>
          </div>
        </div>
      </div>

      <nav className="topnav">
        <div className="topnav-links">
          {NAV_ITEMS.map((item) => (
            <button
              key={item.key}
              className={item.key === current ? "nav-item active" : "nav-item"}
              onClick={() => onNavigate(item.key)}
            >
              {item.label}
            </button>
          ))}
        </div>
      </nav>
    </div>
  );
}
