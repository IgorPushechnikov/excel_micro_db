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
    <main className="min-h-screen flex flex-col items-center justify-center pt-[10vh] bg-gray-100 text-gray-900 dark:text-gray-100 dark:bg-gray-800">
      <h1 className="text-3xl font-bold mb-4 text-center">Welcome to Tauri + React</h1>

      <div className="flex flex-wrap justify-center gap-4 mb-6">
        <a 
          href="https://vite.dev" 
          target="_blank" 
          rel="noopener noreferrer"
          className="font-medium text-blue-600 hover:text-blue-800 dark:hover:text-blue-300 no-underline transition-colors duration-200"
        >
          <img 
            src="/vite.svg" 
            className="h-[6em] w-auto p-[1.5em] will-change-transform transition-[filter] duration-750 hover:drop-shadow-[0_0_2em_#747bff]" 
            alt="Vite logo" 
          />
        </a>
        <a 
          href="https://tauri.app" 
          target="_blank" 
          rel="noopener noreferrer"
          className="font-medium text-blue-600 hover:text-blue-800 dark:hover:text-blue-300 no-underline transition-colors duration-200"
        >
          <img 
            src="/tauri.svg" 
            className="h-[6em] w-auto p-[1.5em] will-change-transform transition-[filter] duration-750 hover:drop-shadow-[0_0_2em_#24c8db]" 
            alt="Tauri logo" 
          />
        </a>
        <a 
          href="https://react.dev" 
          target="_blank" 
          rel="noopener noreferrer"
          className="font-medium text-blue-600 hover:text-blue-800 dark:hover:text-blue-300 no-underline transition-colors duration-200"
        >
          <img 
            src={reactLogo} 
            className="h-[6em] w-auto p-[1.5em] will-change-transform transition-[filter] duration-750 hover:drop-shadow-[0_0_2em_#61dafb]" 
            alt="React logo" 
          />
        </a>
      </div>
      <p className="mb-6 text-center max-w-md">
        Click on the Tauri, Vite, and React logos to learn more.
      </p>

      <form
        className="flex flex-wrap justify-center items-center gap-2 my-4"
        onSubmit={(e) => {
          e.preventDefault();
          greet();
        }}
      >
        <input
          id="greet-input"
          onChange={(e) => setName(e.currentTarget.value)}
          placeholder="Enter a name..."
          className="border border-gray-300 rounded-lg p-2 text-base font-medium text-gray-900 bg-white shadow-[0_2px_2px_rgba(0,0,0,0.2)] focus:outline-none focus:border-blue-500 mr-2 mb-2"
        />
        <button 
          type="submit"
          className="cursor-pointer border border-transparent rounded-lg p-2 text-base font-medium text-white bg-blue-500 shadow-[0_2px_2px_rgba(0,0,0,0.2)] transition-colors duration-200 hover:border-blue-600 active:border-blue-600 active:bg-gray-200 dark:active:bg-gray-800 mb-2"
        >
          Greet
        </button>
      </form>
      <p className="mt-4">{greetMsg}</p>
    </main>
  );
}

export default App;