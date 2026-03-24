import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../authConfig";

export default function Login() {
  const { instance } = useMsal();

  const handleLogin = () => {
    instance.loginPopup(loginRequest).catch(console.error);
  };

  return (
    <div className="login-container">
      <div className="login-card">
        <div className="login-logo">✔</div>
        <h1>Microsoft To-Do</h1>
        <p>Sign in with your Microsoft account to access your tasks.</p>
        <button className="btn btn-primary" onClick={handleLogin}>
          Sign in with Microsoft
        </button>
      </div>
    </div>
  );
}
