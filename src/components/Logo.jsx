import logoIcon from "../assets/logo-icon.png";

export default function Logo({ size = 30 }) {
  return (
    <img
      className="logomark"
      src={logoIcon}
      alt="Timely"
      style={{ height: size, width: "auto", display: "block" }}
    />
  );
}
