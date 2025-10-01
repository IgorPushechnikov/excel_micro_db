// frontend/src/App.tsx
import React from 'react';
import MainLayout from './components/layout/MainLayout'; // Импортируем MainLayout
// Удаляем старые импорты, которые больше не нужны в App.tsx
// import reactLogo from "./assets/react.svg";
// import { invoke } from "@tauri-apps/api/core";

function App() {
  // Удаляем старое состояние и функцию greet
  // const [greetMsg, setGreetMsg] = useState("");
  // const [name, setName] = useState("");
  // async function greet() { ... }

  return (
    // Удаляем старую разметку
    // <main className="container bg-gray-100">...</main>
    <MainLayout /> // Рендерим MainLayout
  );
}

export default App;