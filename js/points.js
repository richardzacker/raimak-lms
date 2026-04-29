const Points = {
  rewards: {
    Sold: 100,
    "Do Not Call": 3,
    FNQ: 3,
    "Pending Order": 3,
    "Already has Fiber": 5,
    TDM: 10,
    "1st Contact": 5,
    "2nd Contact": 10,
    "3rd Contact": 15,
    callbackSet: 10,
    dailyLogin: 5,
  },

  wakeUpHUD() {
    const hud = document.getElementById("economy-hud");
    if (!hud) return;

    // Slide it up!
    hud.classList.remove("hud-hidden");

    // Clear any existing sleep countdowns
    clearTimeout(this._hudSleepTimer);

    // Set a new timer to hide it after 5 seconds of inactivity
    this._hudSleepTimer = setTimeout(() => {
      this._hudSleepTimer = null; // Timer is officially done

      // Only hide it if the user isn't currently hovering over it!
      if (!this._isMouseNearBottom) {
        hud.classList.add("hud-hidden");
      }
    }, 5000);
  },

  initHUDAutoHider() {
    // Wake it up on initial load so they see their balance when they log in
    this.wakeUpHUD();

    // Track the mouse globally (doesn't block clicks like an invisible div would!)
    document.addEventListener("mousemove", (e) => {
      // If the mouse is within the bottom 100 pixels of the screen
      if (e.clientY > window.innerHeight - 100) {
        this._isMouseNearBottom = true;
        const hud = document.getElementById("economy-hud");
        if (hud) hud.classList.remove("hud-hidden");
      } else {
        this._isMouseNearBottom = false;

        // If they move the mouse away, AND the 5-second sleep timer isn't running...
        if (!this._hudSleepTimer) {
          const hud = document.getElementById("economy-hud");
          if (hud) hud.classList.add("hud-hidden");
        }
      }
    });
  },

  async fetchBalances() {
    console.log("🏦 Fetching agent point balances...");

    try {
      // 1. Ask graph.js to go get the raw data
      const allScores = await Graph.getAgentScores();
      State.agentScores = allScores;

      // 2. Figure out who is currently logged in
      const userEmail = ((State.currentUser && State.currentUser.email) || "")
        .toLowerCase()
        .trim();
      const userName =
        (State.currentUser && State.currentUser.name) || userEmail;

      // 3. Find the current user's specific row
      let myScore = allScores.find(
        (s) => (s.AgentEmail || "").toLowerCase().trim() === userEmail,
      );

      // 4. THE INTERCEPT: If they don't exist in the database, build their account!
      if (!myScore && userEmail) {
        console.log(
          `🆕 No point bank found for ${userName}. Creating a new account...`,
        );
        myScore = await Graph.createAgentScore(userEmail, userName);

        // Push this brand new score into the State array so the leaderboard sees them
        State.agentScores.push(myScore);
      }

      // 5. Update our local bank (Using optional chaining fallback just in case)
      this.currentBalance = myScore ? myScore.CurrentPoints : 0;
      this.lifetimePoints = myScore ? myScore.LifetimePoints : 0;

      console.log(
        `💰 Bank Loaded! Balance: ${this.currentBalance} | Lifetime: ${this.lifetimePoints}`,
      );
    } catch (error) {
      console.error("Failed to load point balances:", error);
    }
  },

  getDailySalesCount() {
    const myName = (State.currentUser && State.currentUser.name) || "";
    if (!myName) return 0;

    const uniqueSalesToday = new Set();
    const todayString = new Date().toDateString(); // ⚡ Establish Today

    // 🧽 THE FIX: Give the HUD the Universal Name Cleaner too!
    const cleanName = (str) =>
      (str || "")
        .replace(/[^\w\s]/gi, "")
        .toLowerCase()
        .trim();

    (State.activityLog || []).forEach((log) => {
      // ⚡ Check the date!
      const logDate = new Date(log.timestamp).toDateString();

      if (
        log.action === "Status: " + Config.soldStatus &&
        logDate === todayString
      ) {
        const pastLead = (State.leads || []).find(
          (l) => String(l.id) === String(log.leadId),
        );

        if (pastLead) {
          // ⚡ Compare the scrubbed names!
          const isMine = cleanName(pastLead.assignedTo) === cleanName(myName);
          const isStillSold = pastLead.status === Config.soldStatus;

          if (isMine && isStillSold) {
            uniqueSalesToday.add(log.leadId);
          }
        }
      }
    });

    return uniqueSalesToday.size;
  },

  async awardPoints(actionType, leadId = null) {
    // 1. Validate the action exists in our economy manifest
    let pointValue = this.rewards[actionType];
    if (!pointValue) {
      console.warn(
        `⚠️ Action "${actionType}" is not set in the rewards manifest.`,
      );
      return false;
    }

    // Default to the person currently logged in and clicking the mouse
    let beneficiaryEmail = (
      (State.currentUser && State.currentUser.email) ||
      ""
    )
      .toLowerCase()
      .trim();
    let beneficiaryName = (State.currentUser && State.currentUser.name) || "";

    // 🛡️ 1.5 THE ULTIMATE DRY BOUNCER & BENEFICIARY ROUTER 🛡️
    // Check if the actionType exactly matches a status in our Config array
    if (
      Config.leadStatuses &&
      Config.leadStatuses.includes(actionType) &&
      leadId
    ) {
      const realLead = (State.leads || []).find(
        (l) => String(l.id) === String(leadId),
      );

      // Security Check A: Does the lead exist?
      if (!realLead) {
        console.error(
          "🛑 Security Block: Attempted to claim points for a non-existent lead.",
        );
        return false;
      }

      // Security Check B: Does the actual lead status match the action being claimed?
      if (realLead.status !== actionType) {
        console.error(
          `🛑 Security Block: Lead status is currently '${realLead.status}', but expected '${actionType}'.`,
        );
        return false;
      }

      // --- THE PROXY FIX ---
      // If the lead has an assigned agent, look up their actual email in the Bank
      if (realLead.assignedTo) {
        // HELPER: Strips all punctuation (periods, commas) and extra spaces, then lowercases it
        const normalize = (str) =>
          (str || "")
            .replace(/[^\w\s]/gi, "")
            .replace(/\s+/g, " ")
            .toLowerCase()
            .trim();

        const normalizedLeadOwner = normalize(realLead.assignedTo);

        const ownerBankRow = (State.agentScores || []).find(
          (s) => normalize(s.AgentName) === normalizedLeadOwner,
        );

        if (ownerBankRow && ownerBankRow.AgentEmail) {
          // Switch the beneficiary from the person clicking to the true lead owner
          beneficiaryEmail = ownerBankRow.AgentEmail.toLowerCase().trim();
          beneficiaryName = ownerBankRow.AgentName;
        } else {
          console.error(
            `🛑 System Block: Could not find a bank account for lead owner: ${realLead.assignedTo}.`,
          );
          return false; // Can't pay a ghost!
        }
      }
    }

    // Fallback: If we somehow still don't have an email, abort.
    if (!beneficiaryEmail) return false;

    // 🌟 2. THE COMBO COUNTER LOGIC (Upgraded with Passive Streak Aura) 🌟

    // 1. Calculate the current sales streak no matter WHAT action they just took
    const uniqueSalesToday = new Set();
    const todayString = new Date().toDateString();
    const cleanName = (str) =>
      (str || "")
        .replace(/[^\w\s]/gi, "")
        .toLowerCase()
        .trim();

    (State.activityLog || []).forEach((log) => {
      const logDate = new Date(log.timestamp).toDateString();

      // We only hunt for SALES in the history, excluding the current lead
      if (
        log.action === "Status: " + Config.soldStatus &&
        log.leadId !== leadId &&
        logDate === todayString
      ) {
        const pastLead = (State.leads || []).find(
          (l) => String(l.id) === String(log.leadId),
        );

        if (pastLead) {
          const isMine =
            cleanName(pastLead.assignedTo) === cleanName(beneficiaryName);
          const isStillSold = pastLead.status === Config.soldStatus;

          if (isMine && isStillSold) {
            uniqueSalesToday.add(log.leadId);
          }
        }
      }
    });

    const previousSalesCount = uniqueSalesToday.size;

    // 2. Apply the specific math based on what button they actually clicked
    if (actionType === Config.soldStatus) {
      // 💥 MAIN EVENT: The Sales Multiplier
      const comboMultiplier = previousSalesCount + 1;
      pointValue = pointValue * comboMultiplier;

      if (comboMultiplier > 1) {
        console.log(
          `🔥 COMBO x${comboMultiplier}! Reward boosted to ${pointValue} points!`,
        );
      }
    } else {
      // ⚡ PASSIVE AURA: The Flat Bonus for everything else (+2 XP per past sale)
      if (previousSalesCount > 0) {
        const bonusPoints = previousSalesCount * 2;
        pointValue = pointValue + bonusPoints;

        console.log(
          `⚡ STREAK BONUS! +${bonusPoints} points added to ${actionType} for your ${previousSalesCount}-sale streak!`,
        );
      }
    }

    try {
      // 3. THE BOUNCER: Check for duplicates in the database Ledger
      if (leadId) {
        const isDuplicate = await Graph.checkLedgerForDuplicate(
          leadId,
          actionType,
        );
        if (isDuplicate) {
          console.log(
            `🛑 Bouncer blocked duplicate reward: ${actionType} on lead ${leadId}`,
          );
          return false;
        }
      }

      // 4. THE MATH: Calculate the new totals for the Beneficiary
      const myScoreIndex = State.agentScores.findIndex(
        (s) => s.AgentEmail.toLowerCase() === beneficiaryEmail,
      );

      let myScoreId = null;
      let newCurrentPoints = pointValue;
      let newLifetimePoints = pointValue;

      if (myScoreIndex !== -1) {
        // Add to their existing totals
        newCurrentPoints =
          State.agentScores[myScoreIndex].CurrentPoints + pointValue;
        newLifetimePoints =
          State.agentScores[myScoreIndex].LifetimePoints + pointValue;

        // Update the global State array so the Leaderboard stays live
        State.agentScores[myScoreIndex].CurrentPoints = newCurrentPoints;
        State.agentScores[myScoreIndex].LifetimePoints = newLifetimePoints;
        myScoreId = State.agentScores[myScoreIndex].id;
      }

      const currentUserEmail = (
        (State.currentUser && State.currentUser.email) ||
        ""
      )
        .toLowerCase()
        .trim();
      const isMeClicking = beneficiaryEmail === currentUserEmail;
      // ONLY update the local UI Bank (this.currentBalance) if the person clicking is the beneficiary
      // 🆙 THE LEVEL UP TRACKER 🆙
      const oldLevel = this.calculateLevel(this.lifetimePoints);
      const newLevel = this.calculateLevel(newLifetimePoints);

      if (isMeClicking) {
        this.currentBalance = newCurrentPoints;
        this.lifetimePoints = newLifetimePoints;

        console.log(
          `🎉 Awarded ${pointValue} points for ${actionType}! New Balance: ${this.currentBalance}`,
        );

        // 🚀 TRIGGER THE DOPAMINE (The Decoupling) 🚀
        // We removed this.updateHUD() from here. The HUD will now update
        // exactly when the particles collide in the flyParticlesToHUD function!

        let maxParticles = 12; // Standard burst
        if (pointValue >= 100) maxParticles = 25; // Medium burst
        if (pointValue >= 200) maxParticles = 35; // MASSIVE explosion for Sales!
        if (pointValue >= 300) maxParticles = 40;
        if (pointValue >= 400) maxParticles = 50;

        const minParticles = Math.ceil(maxParticles * 0.25);
        const particleCount =
          Math.floor(Math.random() * (maxParticles - minParticles + 1)) +
          minParticles;

        this.flyParticlesToHUD(particleCount, pointValue);

        // Did we just level up?!
        if (newLevel > oldLevel) {
          console.log(`🎊 LEVEL UP! You are now Level ${newLevel}! 🎊`);

          Points.triggerLevelUp(newLevel);
        }
      } else {
        console.log(
          `🤝 Proxy Sale! Awarded ${pointValue} points to ${beneficiaryName}'s account.`,
        );
        // Note: We don't show the level-up fireworks to the proxy clicker, just log it.
      }

      // 5. THE RECEIPT: Write the transaction with the MULTIPLIED value
      await Graph.writeLedgerTransaction(
        beneficiaryEmail,
        actionType,
        pointValue,
        leadId,
      );

      // 6. THE BANK DEPOSIT: Save the new total to SharePoint
      if (myScoreId) {
        await Graph.updateAgentScore(
          myScoreId,
          newCurrentPoints,
          newLifetimePoints,
        );
      }

      return true;
    } catch (error) {
      console.error("Economy Engine Error:", error);
      return false;
    }
  },

  async purchaseItem(itemId, cost) {
    // Subtract points and log the purchase
  },

  // ==========================================
  // 📈 THE PROGRESSION SYSTEM (XP & LEVELS) 📈
  // ==========================================

  // Calculates the agent's current level based on their Lifetime Points (XP)
  calculateLevel(xp) {
    // Square Root Curve: Divisor of 150. Level 50 takes ~360k XP.
    return Math.floor(Math.sqrt((xp || 0) / 125)) + 1;
  },

  // Calculates exactly how much total XP is required to reach a specific level
  xpForNextLevel(targetLevel) {
    // To find out how much XP is needed to reach Level X
    return Math.pow(targetLevel - 1, 2) * 125;
  },

  triggerLevelUp(newLevel) {
    const banner = document.getElementById("level-up-banner");
    const numSpan = document.getElementById("lvl-up-number");

    if (!banner || !numSpan) return;

    // Update the number
    numSpan.textContent = newLevel;

    // Slam it onto the screen
    banner.classList.add("is-active");

    // Trigger an absolute MAX OVERLOAD of particles (100 is safe for the HD 530!)
    // They will explode out right as the banner slams in.
    this.flyParticlesToHUD(50);

    // Leave it on screen for 3.5 seconds, then fade it out
    setTimeout(() => {
      banner.classList.remove("is-active");
    }, 3500);
  },

  updateHUD(glowDuration = 800) {
    const elLevel = document.getElementById("hud-level");
    const elFill = document.getElementById("hud-fill");
    const elXpCurrent = document.getElementById("hud-xp-current");
    const elXpNext = document.getElementById("hud-xp-next");
    const elBalance = document.getElementById("hud-balance");

    if (!elLevel || !elFill) return; // Fail silently if HUD isn't on the page

    // 1. Calculate the progression stats
    const currentLevel = this.calculateLevel(this.lifetimePoints);
    const xpFloor = this.xpForNextLevel(currentLevel); // XP required for CURRENT level
    const xpCeiling = this.xpForNextLevel(currentLevel + 1); // XP required for NEXT level

    // 2. Math for the Progress Bar
    const xpEarnedInCurrentTier = this.lifetimePoints - xpFloor;
    const xpNeededForNextTier = xpCeiling - xpFloor;
    let fillPercentage = (xpEarnedInCurrentTier / xpNeededForNextTier) * 100;

    // Prevent visual overflow
    if (fillPercentage > 100) fillPercentage = 100;
    if (fillPercentage < 0) fillPercentage = 0;

    // 3. Update the DOM
    const currentDOMLevel = parseInt(elLevel.textContent) || 1;

    elLevel.textContent = currentLevel;
    const oldWidth = elFill.style.width;

    // ⚡ THE SNAP TO ZERO FIX ⚡
    // If the new level is higher than what was just on the screen...
    if (currentLevel > currentDOMLevel) {
      // 1. Turn off the CSS animation completely
      elFill.style.transition = "none";
      // 2. Snap the bar to 0%
      elFill.style.width = "0%";
      // 3. THE MAGIC TRICK: Asking for offsetWidth forces the browser to redraw the screen immediately!
      void elFill.offsetWidth;
      // 4. Turn the CSS animation back on (empty string resets it to your stylesheet rules)
      elFill.style.transition = "";
    }

    // Now set the actual new destination, and it will slide smoothly from 0!
    elFill.style.width = `${fillPercentage}%`;

    // 🌟 TRIGGER THE BAR GLOW 🌟
    if (oldWidth && oldWidth !== elFill.style.width) {
      elFill.classList.add("is-filling");
      clearTimeout(this._glowTimer);

      this._glowTimer = setTimeout(() => {
        if (elFill) elFill.classList.remove("is-filling");
      }, glowDuration);
    }

    // Show exact numbers: e.g., "45 / 375 XP"
    elXpCurrent.textContent = xpEarnedInCurrentTier.toLocaleString();
    elXpNext.textContent = xpNeededForNextTier.toLocaleString();

    // Update the spendable bank
    elBalance.textContent = this.currentBalance.toLocaleString();

    const elComboContainer = document.getElementById("hud-combo-container");
    const elComboCount = document.getElementById("hud-combo-count");

    if (elComboContainer && elComboCount) {
      const currentSales = this.getDailySalesCount();
      const oldSales = parseInt(elComboCount.textContent) || 0;

      // 1. Update the number
      elComboCount.textContent = currentSales;

      // 2. Light it up if they have at least 1 sale
      if (currentSales >= 3) {
        // 🔥 SUPER STREAK: 3+ Sales
        elComboContainer.classList.add("is-hot", "on-fire");
      } else if (currentSales > 0) {
        // ⚡ REGULAR COMBO: 1-2 Sales
        elComboContainer.classList.add("is-hot");
        elComboContainer.classList.remove("on-fire");
      } else {
        // 💤 SLEEPING: 0 Sales
        elComboContainer.classList.remove("is-hot", "on-fire");
      }

      // 3. Trigger the violent bump animation if the number went UP!
      if (currentSales > oldSales) {
        // The Forced Reflow trick to restart the animation instantly
        elComboCount.classList.remove("combo-bump");
        void elComboCount.offsetWidth;
        elComboCount.classList.add("combo-bump");
      }
    }
  },

  // ==========================================
  // ✨ THE DOPAMINE PARTICLE ENGINE (v2.2) ✨
  // ==========================================
  flyParticlesToHUD(particleCount = 5, pointsEarned = 0) {
    const hud = document.getElementById("economy-hud");
    const wasHidden = hud && hud.classList.contains("hud-hidden");
    this.wakeUpHUD();
    const mathDelay = wasHidden ? 350 : 0;

    // Delay the math and the explosion until the HUD is securely locked in place
    setTimeout(() => {
      const targetEl = document.querySelector(".hud-badge");
      if (!targetEl) {
        this.updateHUD(); // Fallback if HUD is broken
        return;
      }

      // 1. Set the Origin (Now guaranteed to be accurate!)
      const targetRect = targetEl.getBoundingClientRect();
      const originX = targetRect.left + targetRect.width / 2;
      const originY = targetRect.top + targetRect.height / 2;

      let maxTotalTime = 0;
      let minTotalTime = Infinity;

      for (let i = 0; i < particleCount; i++) {
        // ⚡ THE FIX: Stagger the spawn by 8 milliseconds per particle
        const spawnDelay = i * 5;

        // --- RANDOMIZED TIMINGS FOR STAGGERING ---
        // (Calculated synchronously so our total times update correctly!)
        const burstDuration = 300 + Math.random() * 300;
        const restDuration = 200 + Math.random() * 500;
        const returnDuration = 400 + Math.random() * 500;

        const totalTime =
          spawnDelay + burstDuration + restDuration + returnDuration;
        if (totalTime > maxTotalTime) maxTotalTime = totalTime;
        if (totalTime < minTotalTime) minTotalTime = totalTime;

        // Delay the actual DOM creation and animation
        setTimeout(() => {
          const particle = document.createElement("div");
          particle.className = "xp-particle";
          document.body.appendChild(particle);

          // Randomize Size (between 6px and 16px)
          const size = Math.random() * 10 + 6;
          particle.style.width = `${size}px`;
          particle.style.height = `${size}px`;

          // Spawn exactly at the center of the HUD badge
          particle.style.left = `${originX - size / 2}px`;
          particle.style.top = `${originY - size / 2}px`;

          particle.style.opacity = "0";
          particle.style.transform = "translate(0, 0) scale(0)";

          // Desynchronize the strobe
          particle.style.animationDelay = `-${Math.random()}s`;

          requestAnimationFrame(() => {
            requestAnimationFrame(() => {
              // PHASE 1: THE BURST
              const minAngle = 205 * (Math.PI / 180);
              const angleRange = 130 * (Math.PI / 180);

              const angle = minAngle + Math.random() * angleRange;

              const distance = 80 + Math.random() * 160;

              const burstX = Math.cos(angle) * distance;
              const burstY = Math.sin(angle) * distance;

              particle.style.transition = `transform ${burstDuration}ms cubic-bezier(0.175, 0.885, 0.32, 1.275), opacity 0.2s ease-out`;
              particle.style.opacity = "1";
              particle.style.transform = `translate(${burstX}px, ${burstY}px) scale(1)`;

              // PHASE 2 & 3: THE REST & RETURN
              setTimeout(() => {
                particle.style.transition = `transform ${returnDuration}ms cubic-bezier(0.55, 0.085, 0.68, 0.53), opacity ${returnDuration}ms ease-in`;
                particle.style.transform = "translate(0px, 0px) scale(0.3)";
                particle.style.opacity = "0.5";

                // PHASE 4: THE IMPACT
                setTimeout(() => {
                  particle.remove();
                  targetEl.style.transform = "scale(1.1)";
                  setTimeout(() => (targetEl.style.transform = "scale(1)"), 50);
                }, returnDuration);
              }, burstDuration + restDuration);
            });
          });
        }, spawnDelay); // <--- Matches the offset for this specific particle!
      }

      // ⚡ PHASE 5: THE BANK DEPOSIT TRIGGER ⚡
      setTimeout(() => {
        const barrageDuration = maxTotalTime - minTotalTime;
        this.updateHUD(Math.max(800, barrageDuration));

        // 💬 FIRE THE FLOATING TEXT!
        this.spawnFloatingText(pointsEarned);
      }, minTotalTime);
    }, mathDelay); // <--- This holds everything back until the HUD stops moving!
  },

  spawnFloatingText(pointsEarned) {
    if (!pointsEarned) return;

    const badge = document.querySelector(".hud-badge");
    if (!badge) return;

    const rect = badge.getBoundingClientRect();
    const text = document.createElement("div");
    text.className = "floating-xp-text";
    text.textContent = `+${pointsEarned}`;

    // Position it slightly above and to the right of the Level Badge
    text.style.left = `${rect.left + 40}px`;
    text.style.top = `${rect.top - 10}px`;

    document.body.appendChild(text);

    // Clean it up after the 1.2s CSS animation finishes
    setTimeout(() => {
      text.remove();
    }, 1200);
  },
};
