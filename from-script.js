document.addEventListener("DOMContentLoaded", function() {
    const saveButton = document.getElementById("saveButton");
    
    // ดึง Element แต่ละช่องผ่าน ID ตรงๆ
    const txtSeqNo = document.getElementById("txtSeqNo");
    const txtFullName = document.getElementById("txtFullName");
    const txtAddress = document.getElementById("txtAddress");
    const txtEmpCode = document.getElementById("txtEmpCode");

    // --- ตารางพักข้อมูล: เก็บรายการที่บันทึกแล้ว ให้บอท RPA มาอ่าน ---
    const STORE_KEY = "accountRecords";
    const recordBody = document.getElementById("recordBody");
    let records = [];
    try { records = JSON.parse(localStorage.getItem(STORE_KEY)) || []; } catch (e) {}

    function renderRecords() {
        recordBody.innerHTML = "";
        records.forEach(function(r) {
            const tr = document.createElement("tr");
            [r.seqNo, r.fullName, r.address, r.empCode].forEach(function(v) {
                const td = document.createElement("td");
                td.textContent = v;
                tr.appendChild(td);
            });
            recordBody.appendChild(tr);
        });
    }

    // --- ป๊อปอัปแบบเลือกได้หลายทาง ใช้ทั้งเตือนรหัสซ้ำ และยืนยันล้างข้อมูล ---
    const actionModal = document.getElementById("actionModal");
    const actionModalMessage = document.getElementById("actionModalMessage");
    const actionModalIcon = document.getElementById("actionModalIcon");
    const actionModalButtons = document.getElementById("actionModalButtons");

    function closeActionModal() {
        if (actionModal) actionModal.classList.remove("show");
    }

    // buttons: [{ label, variant: "primary"|"ghost"|"cancel"|"danger", onClick }]
    function showActionModal(message, icon, buttons) {
        if (!actionModal) return;
        actionModalMessage.textContent = message;
        actionModalIcon.textContent = icon;
        actionModalButtons.innerHTML = "";

        buttons.forEach(function(btn) {
            const b = document.createElement("button");
            b.type = "button";
            b.className = "action-btn action-btn--" + (btn.variant || "ghost");
            b.textContent = btn.label;
            b.addEventListener("click", function() {
                closeActionModal();
                btn.onClick();
            });
            actionModalButtons.appendChild(b);
        });

        actionModal.classList.add("show");
    }

    // existingIndex >= 0 = ทับแถวเดิมตำแหน่งนั้น, -1 = เพิ่มเป็นแถวใหม่
    function commitRecord(rec, existingIndex) {
        if (existingIndex >= 0) { records[existingIndex] = rec; } else { records.push(rec); }
        try { localStorage.setItem(STORE_KEY, JSON.stringify(records)); } catch (e) {}
        renderRecords();
        runSaveAnimation();
    }

    const clearRecordsBtn = document.getElementById("clearRecordsBtn");
    if (clearRecordsBtn) {
        clearRecordsBtn.addEventListener("click", function() {
            if (records.length === 0) return;
            showActionModal(
                "ต้องการล้างรายการทั้งหมด " + records.length + " แถวใช่ไหมคะ? ลบแล้วกู้คืนไม่ได้นะคะ",
                "🗑️",
                [
                    { label: "ล้างข้อมูล", variant: "danger", onClick: function() {
                        records = [];
                        try { localStorage.removeItem(STORE_KEY); } catch (e) {}
                        renderRecords();
                    } },
                    { label: "ยกเลิก", variant: "cancel", onClick: function() {} }
                ]
            );
        });
    }

    renderRecords();

    // --- 1. เอฟเฟกต์พิมพ์ดีดข้อความต้อนรับ ---
    const heading = document.querySelector(".header-section h2");
    if (heading) {
        const text = "กรอกข้อมูลทางบัญชี";
        heading.textContent = "";
        let i = 0;
        function typeWriter() {
            if (i < text.length) {
                heading.textContent += text.charAt(i);
                i++;
                setTimeout(typeWriter, 120);
            }
        }
        setTimeout(typeWriter, 400);
    }

    // --- 2. ระบบสลับธีมกลางคืน / กลางวัน ---
    const themeToggleBtn = document.getElementById("themeToggleBtn");
    if (themeToggleBtn) {
        themeToggleBtn.addEventListener("click", function() {
            document.body.classList.toggle("dark-mode");
            if (document.body.classList.contains("dark-mode")) {
                themeToggleBtn.innerHTML = "☀️ ธีมกลางวัน";
            } else {
                themeToggleBtn.innerHTML = "🌙 เปลี่ยนธีม";
            }
        });
    }

    // --- 3. ตัวควบคุม Popup แจ้งเตือน ---
    const modal = document.getElementById("customModal");
    const modalMessage = document.getElementById("modalMessage");
    const modalIcon = document.getElementById("modalIcon");
    const modalCloseBtn = document.getElementById("modalCloseBtn");

    function showModal(message, icon = "⚠️") {
        modalMessage.textContent = message;
        modalIcon.textContent = icon;
        modal.classList.add("show");
    }

    if (modalCloseBtn) {
        modalCloseBtn.addEventListener("click", function() {
            modal.classList.remove("show");
        });
    }

    

    // --- 5. ระบบตรวจสอบฟอร์มและสถานะปุ่มบันทึก ---

    // แอนิเมชันตอนบันทึกสำเร็จ + ล้างช่องกรอก (ใช้ร่วมกันทั้งบันทึกใหม่ / อัปเดต / )
    // แอนิเมชันตอนบันทึกสำเร็จ + ล้างช่องกรอก (ฉบับลดดีเลย์ ทำงานไวทันใจ)
    function runSaveAnimation() {
        if (typeof window.sparkBurst === "function") {
            const rect = saveButton.getBoundingClientRect();
            window.sparkBurst(rect.left + rect.width / 2, rect.top + rect.height / 2, 20);
        }

        saveButton.classList.add("is-loading");
         
        setTimeout(() => {
            saveButton.classList.remove("is-loading");
            saveButton.classList.add("is-done");

            const label = saveButton.querySelector(".lb-label");
            if (label) {
                label.innerHTML = `<svg class="lb-check" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"></polyline></svg> บันทึกสำเร็จ!`;
            }

            setTimeout(() => {
                if (txtSeqNo) txtSeqNo.value = "";
                if (txtFullName) txtFullName.value = "";
                if (txtAddress) txtAddress.value = "";
                if (txtEmpCode) txtEmpCode.value = "";

                saveButton.classList.remove("is-done");
                if (label) {
                    label.innerHTML = "บันทึกข้อมูล";
                }
            }, 600);

        }, 400);
    }

    if (saveButton) {
        saveButton.addEventListener("click", function(event) {
            event.preventDefault();

            // เช็กช่องรหัส
            if (txtSeqNo && txtSeqNo.value.trim() === "") {
                showModal("กรุณากรอกรหัสก่อนทำการบันทึกค่ะ!", "⚠️");
                txtSeqNo.focus();
                return;
            }

            // ตัดช่องว่าง + รหัสเป็นตัวพิมพ์ใหญ่ เพื่อให้บอทเทียบกับ Excel ได้ตรง
            const rec = {
                seqNo: txtSeqNo.value.trim().toUpperCase(),
                fullName: txtFullName.value.trim(),
                address: txtAddress.value.trim(),
                empCode: txtEmpCode.value.trim()
            };

            const dupIndex = records.findIndex(function(r) { return r.seqNo === rec.seqNo; });

            if (dupIndex === -1) {
                // รหัสยังไม่เคยมีในตาราง บันทึกเป็นแถวใหม่ได้เลย
                commitRecord(rec, -1);
                return;
            }

            // รหัสซ้ำ: ถามก่อนว่าจะอัปเดตแถวเดิม หรือยกเลิก 
            // ยกเลิก = ไม่บันทึกอะไรเลย ข้อมูลในฟอร์มยังอยู่ครบ แก้รหัสแล้วกดบันทึกใหม่ได้ (เหมือนตอนเตือนว่าลืมกรอกรหัส)
            // ถ้าเจอ รหัสซ้ำ ให้ถามแค่อัปเดต หรือ ยกเลิก (ตัด "เพิ่มเป็นแถวใหม่" ออกไปเลย)
            showActionModal(
                "พบรหัส \"" + rec.seqNo + "\" ในระบบแล้วค่ะ ต้องการอัปเดตข้อมูลเดิมใช่ไหมคะ?",
                "🔄",
                [
                    { label: "อัปเดตข้อมูล", variant: "primary", onClick: function() { commitRecord(rec, dupIndex); } },
                    { label: "ยกเลิก/แก้ไขรหัส", variant: "cancel", onClick: function() {} }
                ]
            );
            
        });
    }
});

/* --- พื้นหลังวิบวับ — ดาวกระพริบ + ผงตามเมาส์ --- */
(() => {
  "use strict";

  const CONFIG = { density: 6000, dustDelay: 30, dustLife: 1000 };
  const TINTS = ["124,108,240", "240,165,200", "140,190,240", "250,205,140"];
  const SPARKS = ["✦", "✧", "⋆", "✨"];
  const COLORS = ["#a99cff", "#f0a5c8", "#8cbef0", "#fad08c"];
  const reduced = matchMedia("(prefers-reduced-motion: reduce)").matches;

  const cv = document.getElementById("sky");
  if (cv && !reduced) {
    const ctx = cv.getContext("2d");
    let stars = [], W = 0, H = 0;

    function resize() {
      W = cv.width = innerWidth;
      H = cv.height = innerHeight;
      const count = Math.floor((W * H) / CONFIG.density);
      stars = Array.from({ length: count }, () => ({
        x: Math.random() * W, y: Math.random() * H,
        r: Math.random() * 1.6 + 0.5, a: Math.random() * Math.PI * 2,
        sp: Math.random() * 0.018 + 0.006, dy: Math.random() * 0.14 + 0.03,
        t: TINTS[Math.floor(Math.random() * TINTS.length)]
      }));
    }

    function draw() {
      ctx.clearRect(0, 0, W, H);
      for (const s of stars) {
        s.a += s.sp;
        const alpha = ((Math.sin(s.a) + 1) / 2) * 0.55 + 0.08;
        s.y -= s.dy;
        if (s.y < -5) { s.y = H + 5; s.x = Math.random() * W; }

        ctx.beginPath();
        ctx.arc(s.x, s.y, s.r, 0, Math.PI * 2);
        ctx.fillStyle = `rgba(${s.t},${alpha})`;
        ctx.shadowBlur = s.r * 5;
        ctx.shadowColor = `rgba(${s.t},.5)`;
        ctx.fill();
      }
      ctx.shadowBlur = 0;
      requestAnimationFrame(draw);
    }

    addEventListener("resize", resize);
    resize();
    draw();
  }

  if (!reduced) {
    let last = 0;
    addEventListener("pointermove", (e) => {
      const now = Date.now();
      if (now - last < CONFIG.dustDelay) return;
      last = now;
      const d = document.createElement("div");
      d.className = "dust";
      d.textContent = SPARKS[Math.floor(Math.random() * SPARKS.length)];
      d.style.left = e.clientX + "px";
      d.style.top = e.clientY + "px";
      d.style.color = COLORS[Math.floor(Math.random() * COLORS.length)];
      d.style.setProperty("--dx", (Math.random() * 46 - 23) + "px");
      d.style.setProperty("--dy", (Math.random() * 38 + 14) + "px");
      document.body.appendChild(d);
      setTimeout(() => d.remove(), CONFIG.dustLife);
    });
  }

  window.sparkBurst = function (x, y, amount = 24) {
    if (reduced) return;
    for (let i = 0; i < amount; i++) {
      const p = document.createElement("div");
      p.className = "dust";
      p.textContent = SPARKS[i % SPARKS.length];
      p.style.left = x + "px"; p.style.top = y + "px";
      p.style.fontSize = (10 + Math.random() * 12) + "px";
      p.style.color = COLORS[i % COLORS.length];
      const ang = (Math.PI * 2 * i) / amount;
      const dist = 70 + Math.random() * 90;
      p.style.setProperty("--dx", Math.cos(ang) * dist + "px");
      p.style.setProperty("--dy", Math.sin(ang) * dist + "px");
      p.style.animationDuration = "1.1s";
      document.body.appendChild(p);
      setTimeout(() => p.remove(), 1150);
    }
  };
})();