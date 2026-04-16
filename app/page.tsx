"use client";

import styles from "./page.module.css";
import GenerateForms from "./Components/GenerateForms/GenerateForms";

export default function Home() {
  return (
    <div className={styles.container}>
      <h1 className={styles.title}>Bonus System Dashboard</h1>

      <div className={styles.logoBar}>
        <img
          src="/foundationfirstvalogo.png"
          alt="Foundation First VA Logo"
          className={styles.logo}
        />

        <img
          src="/clear-automate-logo.png"
          alt="Clear Automate Logo"
          className={styles.logo}
        />
      </div>

      {/* Content Area */}
      <div className={styles.contentCard}>
        <div>
          <h2>Generate Forms</h2>
          <GenerateForms />
        </div>
      </div>
    </div>
  );
}