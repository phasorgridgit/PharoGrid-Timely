import { useEffect, useState } from "react";

export default function UpdateBanner() {
  const [status, setStatus] = useState(null); // null | "downloading" | "ready" | "error"
  const [detail, setDetail] = useState({});

  useEffect(() => {
    if (!window.workhub?.onUpdateStatus) return; // mock mode / no Electron bridge
    return window.workhub.onUpdateStatus(({ status: s, ...rest }) => {
      // "checking" and "up-to-date" are intentionally not shown — nothing
      // for the employee to act on, and surfacing every background check
      // would just be noise in an app people leave open all day.
      if (s === "checking" || s === "up-to-date") return;
      setStatus(s);
      setDetail(rest);
    });
  }, []);

  if (!status) return null;

  if (status === "downloading") {
    return (
      <div className="update-banner update-banner-info">
        Downloading update{detail.percent != null ? ` — ${detail.percent}%` : "…"}
      </div>
    );
  }

  if (status === "ready") {
    return (
      <div className="update-banner update-banner-ready">
        <span>Update {detail.version ? `to v${detail.version} ` : ""}ready to install.</span>
        <button onClick={() => window.workhub.quitAndInstallUpdate()}>Restart to update</button>
      </div>
    );
  }

  // "error" — logged for support, not shown to the employee; a failed
  // background update check isn't something they can do anything about, and
  // the app keeps working fine on the current version either way.
  if (status === "error") {
    console.error("Timely update check failed:", detail.message);
    return null;
  }

  return null;
}
