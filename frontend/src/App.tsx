import { useState } from "react";
import reactLogo from "./assets/react.svg";
import { invoke } from "@tauri-apps/api/core";

// Используем Tailwind CSS, импорт App.css не требуется

function App() {
  const [greetMsg, setGreetMsg] = useState("");
  const [name, setName] = useState("");

  async function greet() {
    // Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
    setGreetMsg(await invoke("greet", { name }));
  }

  return (
    <main className="min-h-screen bg-gray-50 p-8">
      <div className="max-w-4xl mx-auto bg-white rounded-lg shadow-md p-6">
        <h1 className="text-2xl font-bold text-blue-600 mb-4">Welcome to Excel Micro DB</h1>
        
        <div className="flex flex-wrap gap-4 mb-6">
          <a href="https://vite.dev" target="_blank" className="hover:opacity-80 transition-opacity">
            <img src="/vite.svg" className="h-12" alt="Vite logo" />
          </a>
          <a href="https://tauri.app" target="_blank" className="hover:opacity-80 transition-opacity">
            <img src="/tauri.svg" className="h-12" alt="Tauri logo" />
          </a>
          <a href="https://react.dev" target="_blank" className="hover:opacity-80 transition-opacity">
            <img src={reactLogo} className="h-12" alt="React logo" />
          </a>
        </div>
        
        <p className="text-gray-700 mb-6">Click on the logos to learn more. Tailwind CSS is working!</p>

        <form
          className="flex flex-col gap-4 mb-6"
          onSubmit={(e) => {
            e.preventDefault();
            greet();
          }}
        >
          <input
            id="greet-input"
            onChange={(e) => setName(e.currentTarget.value)}
            placeholder="Enter a name..."
            className="px-4 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          <button 
            type="submit" 
            className="px-6 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors font-medium"
          >
            Greet
          </button>
        </form>
        
        {greetMsg && (
          <div className="p-4 bg-green-100 text-green-800 rounded-md">
            {greetMsg}
          </div>
        )}
      </div>
    </main>
  );
}

export default App;
