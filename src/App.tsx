import {
  Providers,
  SimpleProvider,
  ProviderState,
} from "@microsoft/mgt-element";
import { Login } from "@microsoft/mgt-react";
import Persons from "./Persons";
import "./App.css";

Providers.globalProvider = new SimpleProvider(
  (_scopes: string[]): Promise<string> => {
    return new Promise((resolve) => {
      resolve("<access_token>");
    });
  }
);

Providers.globalProvider.setState(ProviderState.SignedIn);

function App() {
  return (
    <div className="App">
      <header>
        <Login />
      </header>
      <Persons />
    </div>
  );
}

export default App;
