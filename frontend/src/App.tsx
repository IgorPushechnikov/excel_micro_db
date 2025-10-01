// src/App.tsx
import { useState } from "react";
import reactLogo from "./assets/react.svg";
import { invoke } from "@tauri-apps/api/core";

function App() {
  const [greetMsg, setGreetMsg] = useState("");
  const [name, setName] = useState("");

  async function greet() {
    // Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
    setGreetMsg(await invoke("greet", { name }));
  }

  return (
    // Добавлен утилитарный класс bg-gray-100, чтобы Tailwind его обнаружил
    <main className="container bg-gray-100">
      <h1>Welcome to Tauri + React</h1>

      <div className="row">
        {/* Добавлен атрибут rel для безопасности и производительности */}
        <a href="https://vite.dev" target="_blank" rel="noopener noreferrer">
          <img src="/vite.svg" className="logo vite" alt="Vite logo" />
        </a>
        {/* Добавлен атрибут rel для безопасности и производительности */}
        <a href="https://tauri.app" target="_blank" rel="noopener noreferrer">
          <img src="/tauri.svg" className="logo tauri" alt="Tauri logo" />
        </a>
        {/* Добавлен атрибут rel для безопасности и производительности */}
        <a href="https://react.dev" target="_blank" rel="noopener noreferrer">
          <img src={reactLogo} className="logo react" alt="React logo" />
        </a>
      </div>
      <p>Click on the Tauri, Vite, and React logos to learn more.</p>

      <form
        className="row"
        onSubmit={(e) => {
          e.preventDefault();
          greet();
        }}
      >
        <input
          id="greet-input"
          onChange={(e) => setName(e.currentTarget.value)}
          placeholder="Enter a name..."
        />
        <button type="submit">Greet</button>
      </form>
      <p>{greetMsg}</p>
    </main>
  );
}

export default App;