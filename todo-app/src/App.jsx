import { MsalProvider, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "./authConfig";
import Login from "./components/Login";
import TaskList from "./components/TaskList";
import "./App.css";

const msalInstance = new PublicClientApplication(msalConfig);

export default function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <AuthenticatedTemplate>
        <TaskList />
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <Login />
      </UnauthenticatedTemplate>
    </MsalProvider>
  );
}
