"use client";

import { useState } from "react";
import styles from "./page.module.css";
import GenerateForms from "./Components/GenerateForms/GenerateForms";
import ImportScores from "./Components/ImportScores/ImportScores";

export default function Home() {
  const [activeTab, setActiveTab] = useState<"generate" | "import">("generate");

  return (
    <div className={styles.container}>
      <h1 className={styles.title}>Bonus System Dashboard</h1>

      {/* Tabs */}
      <div className={styles.tabBar}>
        <button
          onClick={() => setActiveTab("generate")}
          className={`${styles.tabButton} ${activeTab === "generate" ? styles.activeTab : ""
            }`}
        >
          Generate Forms
        </button>

        <button
          onClick={() => setActiveTab("import")}
          className={`${styles.tabButton} ${activeTab === "import" ? styles.activeTab : ""
            }`}
        >
          Import Scores
        </button>
      </div>

      {/* Content Area */}
      <div className={styles.contentCard}>
        {activeTab === "generate" && (
          <div>
            <h2>Generate Forms</h2>
            <GenerateForms />
          </div>
        )}

        {activeTab === "import" && (
          <div>
            <h2>Import Scores</h2>
            <ImportScores />
          </div>
        )}
      </div>
    </div>
  );
}
