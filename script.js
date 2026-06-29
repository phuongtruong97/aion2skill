// --- GLOBAL DATA ---
let dbDmg={}, dbAbn={}, dbFilter={}, iconDb={};
let dbDmgLv={}, dbAbnLv={};
let dbSkillLv={}; // THÊM BIẾN NÀY ĐỂ CHỨA DATA CẤP ĐỘ MP/CD
let skillMap={}, iconMap={};
let currentSkill=null, selectedRunes=[], currentLevel=1;
let currentLang = 'en'; 
const CLASS_CODE = {"ASSASSIN":"as","CLERIC":"cl","WARRIOR":"wa","MAGE":"ma","ARCHER":"ar","WARLORD":"wl","BLADEDANCER":"bd","GUNSLINGER":"gs","CHANTER":"ch","TEMPLAR":"te","GLADIATOR":"gl","SORCERER":"so","SPIRITMASTER":"sm","RANGER":"ra"};

// --- 1. AUTO LOAD DATA (UPDATED FOR DYNAMIC FOLDERS) ---
const log = m => document.getElementById('status').innerText = m;
const setStatusError = (msg) => {
    const el = document.getElementById('status');
    el.innerText = msg;
    el.className = 'status-error';
    document.getElementById('ui-spinner').style.borderTopColor = 'red';
    document.getElementById('ui-spinner').style.animation = 'none';
}

// Biến lưu patch hiện tại đang dùng cho trang Tra Cứu
let currentPatchFolder = ""; 

window.onload = async () => {
    if (window.location.protocol === 'file:') {
        setStatusError("LỖI: Chặn bảo mật (CORS)");
        alert("Lỗi file://. Hãy dùng Live Server.");
        return;
    }

    // BƯỚC 1: Đọc file versions.json để lấy Patch mới nhất
    try {
        const vRes = await fetch('./data/versions.json');
        if (!vRes.ok) throw new Error("Không tìm thấy data/versions.json");
        const versions = await vRes.json();
        
        if (versions.length > 0) {
            // Lấy patch đầu tiên trong danh sách làm mặc định
            currentPatchFolder = versions[0].id;
            console.log("Loading latest patch:", currentPatchFolder);
            await loadAllData(currentPatchFolder);
        } else {
            throw new Error("File versions.json trống!");
        }
    } catch (e) {
        console.error(e);
        setStatusError("Lỗi cấu trúc Data: " + e.message);
    }
};

async function loadAllData(folderName) {
    // Đường dẫn trỏ vào folder con (VD: ./data/2025_12_31/Skill.json)
    // Dùng data_min (bản đã lọc field, nhẹ hơn ~86%) thay cho data gốc của game
    const basePath = `./data_min/${folderName}`; 

    const jsonFiles = [
        { name: 'Skill.json', db: iconDb, isSkill: true },
        { name: 'SkillEffect.json', db: dbDmg, isLv: false },
        { name: 'SkillLv.json', db: dbSkillLv, isSkillLv: true },
        { name: 'SkillAbnormalEffect.json', db: dbAbn, isLv: false }, 
        { name: 'SkillEffectLv.json', db: dbDmgLv, isLv: true, isAbn: false },
        { name: 'SkillAbnormalEffectLv.json', db: dbAbnLv, isLv: true, isAbn: true },
        { name: 'SkillEffectFilter.json', db: dbFilter, isLv: false }
    ];

    let loadedCount = 0;
    const updateProgress = () => {
        loadedCount++;
        log(`Loading [${folderName}]: ${loadedCount}/${jsonFiles.length + 1}...`); 
    };

    const jsonPromises = jsonFiles.map(async (f) => {
        try {
            const res = await fetch(`${basePath}/${f.name}`);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const json = await res.json();
            
            if (f.isSkill) scanSkill(json);
            else if (f.isSkillLv) scanSkillLvData(json, f.db);
            else if (f.isLv) scanJsonLv(json, f.db, f.isAbn);
            else scanJson(json, f.db);
            
            updateProgress();
        } catch (e) {
            console.error(`Lỗi tải ${f.name}:`, e);
            setStatusError(`Thiếu file: ${folderName}/${f.name}`);
        }
    });

    const excelPromise = (async () => {
        try {
            const res = await fetch(`${basePath}/text.xlsx`);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const buf = await res.arrayBuffer();
            const wb = XLSX.read(new Uint8Array(buf), {type:'array'});
            processExcel(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1}));
            updateProgress();
        } catch (e) {
            console.error("Lỗi tải Excel:", e);
            setStatusError(`Thiếu file: ${folderName}/text.xlsx`);
        }
    })();

    await Promise.all([...jsonPromises, excelPromise]);

    if (document.getElementById('status').className !== 'status-error') {
        document.getElementById('ui-spinner').style.display = 'none';
        const st = document.getElementById('status');
        st.innerText = `${folderName}`; // Hiện tên Patch lên status
        st.classList.add('status-ready');
        changeLanguage(); 
    }
}

function changeLanguage() {
    currentLang = document.getElementById('ui-lang').value;
    
    // Dịch các nhãn text cơ bản
    document.getElementById('cat-active').innerText = currentLang === 'vi' ? 'Kích Hoạt' : 'Active';
    document.getElementById('cat-stigma').innerText = 'Stigma';
    document.getElementById('cat-passive').innerText = currentLang === 'vi' ? 'Bị Động' : 'Passive';
    document.getElementById('label-spec').innerText = currentLang === 'vi' ? 'Chuyên Hóa' : 'Specialization';
    
    // --- THÊM: Dịch nút Tab ---
    document.getElementById('btn-tab-search').innerText = currentLang === 'vi' ? 'Skill' : 'Skill';
    document.getElementById('btn-tab-compare').innerText = currentLang === 'vi' ? 'Patch Notes' : 'Patch Notes';
    // -------------------------

    if(currentSkill) {
        loadSkill(currentSkill, document.querySelector('.skill-btn.selected'), false);
    }
}

// --- 2. EXCEL PROCESS ---
function processExcel(rows) {
    skillMap = {};
    const regex = /SkillString_STR_SKILL_PC_([A-Z_]+)_(\d{8})_(.+)/;
    let rawById = {};
    
    rows.forEach(r => {
        if(typeof r[0] !== 'string') return;
        const m = r[0].match(regex);
        if(m) {
            const cls = m[1], fid = m[2], type = m[3];
            if(!rawById[cls]) rawById[cls] = {};
            if(!rawById[cls][fid]) rawById[cls][fid] = { id: fid };
            
            const txtEn = r[1] || "";
            const txtVi = r[2] || r[1] || ""; 

            if(type === 'skill_name') {
                rawById[cls][fid].name_en = txtEn;
                rawById[cls][fid].name_vi = txtVi;
            }
            else if(type === 'skill_desc_effect') {
                rawById[cls][fid].desc_en = txtEn;
                rawById[cls][fid].desc_vi = txtVi;
            }
            else if(type === 'skill_spec_effect') {
                rawById[cls][fid].tag_en = txtEn;
                rawById[cls][fid].tag_vi = txtVi;
            }
            else if(type === 'specialized_skill_desc') {
                rawById[cls][fid].runeDesc_en = txtEn;
                rawById[cls][fid].runeDesc_vi = txtVi;
            }
        }
    });

    for(let cls in rawById) {
        skillMap[cls] = { active: [], stigma: [], passive: [] };
        let grouped = {};
        for(let fid in rawById[cls]) {
            const prefix = fid.substring(0, 4); 
            if(!grouped[prefix]) {
                grouped[prefix] = { 
                    prefix: prefix, 
                    baseId: fid, 
                    className: cls, 
                    name_en: rawById[cls][fid].name_en || "Unknown", 
                    name_vi: rawById[cls][fid].name_vi || "Unknown",
                    tag_en: rawById[cls][fid].tag_en || "", 
                    tag_vi: rawById[cls][fid].tag_vi || "",
                    variants: {}, 
                    runes: {}, 
                    hasCombo: false,
                    skillType: null 
                };
                
                const idNum = parseInt(prefix + "0000"); 
                if(iconDb[idNum] && iconDb[idNum].type) {
                    grouped[prefix].skillType = iconDb[idNum].type;
                }
            }
            const g = grouped[prefix];
            const suffix = fid.substring(4);
            
            g.variants[suffix] = {
                desc_en: rawById[cls][fid].desc_en,
                desc_vi: rawById[cls][fid].desc_vi
            };

            if((rawById[cls][fid].runeDesc_en || rawById[cls][fid].runeDesc_vi) && suffix.endsWith('0') && suffix.startsWith('00') && suffix[2] !== '0') {
                g.runes[parseInt(suffix[2])] = {
                    en: rawById[cls][fid].runeDesc_en,
                    vi: rawById[cls][fid].runeDesc_vi
                };
            }

            if(!suffix.startsWith('00') && suffix.endsWith('0')) g.hasCombo = true;
            
            if(suffix === '0000') { 
                if(rawById[cls][fid].name_en) g.name_en = rawById[cls][fid].name_en; 
                if(rawById[cls][fid].name_vi) g.name_vi = rawById[cls][fid].name_vi;
                if(rawById[cls][fid].tag_en) g.tag_en = rawById[cls][fid].tag_en; 
                if(rawById[cls][fid].tag_vi) g.tag_vi = rawById[cls][fid].tag_vi;
            }
        }
        
        for(let key in grouped) {
            const skill = grouped[key];

            // 1. Lấy tên skill và chuyển về chữ thường hết
            const nCheck = (skill.name_en || "").toLowerCase().trim();

            // 2. Danh sách các tên muốn chặn (LƯU Ý: Phải viết chữ thường toàn bộ ở đây)
            const ignoreList = [
                "dodge",
                "basic attack",
                "provoke <dnt>",
                "punishing benediction (không sử dụng)",
                "punishing benediction (not used)",
                "surging bloodlust",
                "bone-chilling roar",
                "curse: old tree",
                "firebomb",
                "vaizel's wisdom",
                "summon flame",
                "illusion",
                "magic energy blast",
                "lumiel's authority",
                "eye of detection",
                "eye of rapid burst",
                "afterimage",
                "hunters resolve_active",
                "arrow of space-time",
                "dust arrow",
                "impact kick",
                "lightning arrow",
                "intimidating roar",
                "rush",
                "doom advent",
                "madness blow",
                "wrathful strike",
                "curse of despair",
                "blessing of regeneration",
                "crushing strike",
                "gladiator weapon equipment",
                "templar weapon equipment",
                "sorcerer weapon equip",
                "chanter_equip weapon",
                "cleric weapon equip",   // Ví dụ thêm
                "assassin weapon equip"  // Ví dụ thêm
            ];

            // 3. Kiểm tra:
            // - Cách A: Nếu tên skill nằm TRONG danh sách chặn -> Bỏ qua
            // - Cách B: Hoặc nếu tên skill CÓ CHỨA cụm từ "weapon equipment" -> Bỏ qua (đỡ phải liệt kê từng class)
            if (ignoreList.includes(nCheck) || nCheck.includes("weapon equipment") || nCheck.includes("equip weapon")) {
                continue; 
            }
            
            // ... (Code cũ giữ nguyên từ đây)
            if(skill.skillType === "ESkillType::Passive") {
                skillMap[cls].passive.push(skill);
            } 
            else if(skill.skillType === "ESkillType::Active") {
            // ... (đoạn code phía dưới giữ nguyên)
                if (Object.keys(skill.runes).length > 0 && !skill.hasCombo) {
                        skillMap[cls].stigma.push(skill);
                } else {
                        skillMap[cls].active.push(skill);
                }
            }
            else {
                if(Object.keys(skill.runes).length === 0) skillMap[cls].passive.push(skill);
                else if (skill.hasCombo) skillMap[cls].active.push(skill);
                else skillMap[cls].stigma.push(skill);
            }
        }
    }
    const sel = document.getElementById('ui-class');
    sel.innerHTML = '<option>-- Select Class --</option>';
    Object.keys(skillMap).sort().forEach(c => sel.innerHTML += `<option value="${c}">${c}</option>`);
    
    changeLanguage();
}

// --- 3. UI RENDER ---
function renderSkillList() {
    const cls = document.getElementById('ui-class').value;
    const gActive = document.getElementById('grid-active'), gStigma = document.getElementById('grid-stigma'), gPassive = document.getElementById('grid-passive');
    gActive.innerHTML = ""; gStigma.innerHTML = ""; gPassive.innerHTML = "";
    if(!skillMap[cls]) return;
    
    const makeBtn = (s) => {
        const d = document.createElement('div'); d.className = 'skill-btn';
        const icon = getIcon(s.className, s.baseId);
        if(icon) d.style.backgroundImage = `url('${icon}')`;
        d.innerHTML = `<div class="id-label">${s.prefix}</div>`;
        d.onclick = () => loadSkill(s, d, true); // true = open overlay
        return d;
    };
    
    skillMap[cls].active.forEach(s => gActive.appendChild(makeBtn(s)));
    skillMap[cls].stigma.forEach(s => gStigma.appendChild(makeBtn(s)));
    skillMap[cls].passive.forEach(s => gPassive.appendChild(makeBtn(s)));
}

function updateLevel(val) { 
    currentLevel = parseInt(val); 
    document.getElementById('lv-display').innerText = currentLevel; 
    
    if(currentSkill) {
        const baseIdNum = parseInt(currentSkill.baseId.replace(/[^0-9]/g, ''));
        const meta = getMeta(baseIdNum, currentLevel); // HÀM GETMETA ĐÃ ĐƯỢC CẬP NHẬT
        document.getElementById('ui-mp').innerText = meta.mp > 0 ? fmt(meta.mp) : "--";
        document.getElementById('ui-cd').innerText = meta.cd > 0 ? `${meta.cd}s` : "--";

        if(currentSkill.type === 'stigma') {
            selectedRunes = [];
            if(currentLevel >= 5) selectedRunes.push(1);
            if(currentLevel >= 10) selectedRunes.push(2);
            if(currentLevel >= 15) selectedRunes.push(3);
            if(currentLevel >= 20) selectedRunes.push(4);
        }
        updateUI();
    }
}

function getRuneHtml(index, isStigma) {
    let imgName = "";
    if (isStigma) {
        imgName = "icon_specialized_skill_stigma_common_001.png";
    } else {
        const num = (index % 5) + 1;
        imgName = `icon_item_usable_stigma_el_a_c_00${num}.png`;
    }
    // LOCAL PATH: ./icons/
    const url = `./icons/${imgName.toLowerCase()}`;
    return `<div class="rune-img" style="background-image: url('${url}')"></div>`;
}

// --- NEW: MOBILE OVERLAY LOGIC ---
function openMobileOverlay() {
    if (window.innerWidth <= 1080) {
        document.getElementById('ui-tooltip').classList.add('mobile-active');
    }
}

function closeMobileOverlay() {
    document.getElementById('ui-tooltip').classList.remove('mobile-active');
}

function loadSkill(skill, btn, shouldOpenOverlay) {
    currentSkill = skill; selectedRunes = [];
    if(btn) {
        document.querySelectorAll('.skill-btn').forEach(b => b.classList.remove('selected'));
        btn.classList.add('selected');
    }
    
    document.getElementById('ui-name').innerText = skill['name_' + currentLang];
    const icon = getIcon(skill.className, skill.baseId);
    document.getElementById('ui-icon').style.backgroundImage = icon ? `url('${icon}')` : 'none';
    document.getElementById('ui-spec-tag').innerHTML = parseTags(skill['tag_' + currentLang]);

    // RESET LEVEL
    const slider = document.getElementById('lv-slider');
    slider.value = 1; 
    currentLevel = 1; 
    document.getElementById('lv-display').innerText = 1;

    const baseIdNum = parseInt(skill.baseId.replace(/[^0-9]/g, ''));
    const meta = getMeta(baseIdNum, 1);
    document.getElementById('ui-mp').innerText = meta.mp > 0 ? fmt(meta.mp) : "--";
    document.getElementById('ui-cd').innerText = meta.cd > 0 ? `${meta.cd}s` : "--";

    const specCont = document.getElementById('ui-spec-container');
    const slotsDiv = document.getElementById('ui-slots');
    const runesDiv = document.getElementById('ui-runes');
    
    slotsDiv.innerHTML = ""; runesDiv.innerHTML = "";
    const runeKeys = Object.keys(skill.runes).map(Number).sort();
    
    if(skill.skillType === "ESkillType::Passive" || runeKeys.length === 0) {
            if(runeKeys.length > 0) {
                    specCont.style.display = "block";
                    slotsDiv.style.display = 'none';
                    document.getElementById('ui-spec-count').innerText = "";
                    currentSkill.type = "passive";
            } else {
                    specCont.style.display = "none"; 
                    currentSkill.type = "passive";
            }
    } 
    else {
        specCont.style.display = "block";
        const isStigma = !skill.hasCombo;
        currentSkill.type = isStigma ? "stigma" : "active";

        if (isStigma) {
            slotsDiv.style.display = 'none';
            document.getElementById('ui-spec-count').innerText = ""; 
        } else {
            slotsDiv.style.display = 'flex';
            const maxSlots = 3;
            document.getElementById('ui-spec-count').innerText = `0/${maxSlots}`;
            for(let i=0; i<maxSlots; i++) {
                const slot = document.createElement('div'); slot.className = 'slot'; slot.id = `slot-${i}`;
                slot.onclick = () => removeRune(i);
                slot.innerHTML = `<div class="slot-remove">✕</div>`;
                slotsDiv.appendChild(slot);
            }
        }
    }

    runeKeys.forEach((idx, i) => {
        const row = document.createElement('div'); row.className = 'rune-row'; row.id = `rune-row-${idx}`;
        row.onclick = () => toggleRune(idx, (currentSkill.type === 'active' ? 3 : 99));
        const runeTxt = skill.runes[idx][currentLang] || skill.runes[idx]['vi'] || "";
        const isStig = currentSkill.type === 'stigma';
        row.innerHTML = `${getRuneHtml(i, isStig)}<div class="rune-text">${parseTags(runeTxt)}</div>`;
        runesDiv.appendChild(row);
    });
    
    if(currentSkill.type === 'stigma') updateLevel(1);
    
    updateUI();

    // Trigger Overlay
    if(shouldOpenOverlay) openMobileOverlay();
}

function toggleRune(idx, max) {
    if (currentSkill.type === 'stigma') {
        let targetLevel = idx * 5; 
        if (currentLevel === targetLevel) {
            targetLevel -= 5;
            if (targetLevel < 1) targetLevel = 1;
        }
        const slider = document.getElementById('lv-slider');
        slider.value = targetLevel;
        updateLevel(targetLevel);
    } else {
        if(selectedRunes.includes(idx)) selectedRunes = selectedRunes.filter(i => i!==idx);
        else { if(selectedRunes.length < max) selectedRunes.push(idx); else alert(`Max ${max}!`); }
        selectedRunes.sort((a,b)=>a-b); 
        updateUI();
    }
}

function removeRune(idx) { 
    if(idx < selectedRunes.length) { 
        selectedRunes.splice(idx, 1); 
        updateUI(); 
    } 
}

function updateUI() {
    if(!currentSkill) return;
    const isStigma = currentSkill.type === 'stigma';
    
    if (currentSkill.type === 'active') {
        const slots = document.querySelectorAll('.slot');
        slots.forEach(s => { s.className = 'slot'; s.innerHTML = '<div class="slot-remove">✕</div>'; });
        
        selectedRunes.forEach((rid, i) => {
            if(slots[i]) {
                slots[i].classList.add('filled');
                const keys = Object.keys(currentSkill.runes).map(Number).sort();
                const runeIndex = keys.indexOf(rid);
                slots[i].innerHTML += getRuneHtml(runeIndex, false);
            }
        });
        document.getElementById('ui-spec-count').innerText = `${selectedRunes.length}/3`;
    }

    document.querySelectorAll('.rune-row').forEach(r => r.classList.remove('active'));
    selectedRunes.forEach(rid => { if(document.getElementById(`rune-row-${rid}`)) document.getElementById(`rune-row-${rid}`).classList.add('active'); });
    
    let suffix = "0000";
    if(selectedRunes.length > 0) {
        if(isStigma) suffix = `00${selectedRunes.length}0`;
        else suffix = (selectedRunes.length===1 ? `00${selectedRunes[0]}0` : (selectedRunes.length===2 ? `0${selectedRunes[0]}${selectedRunes[1]}0` : `${selectedRunes[0]}${selectedRunes[1]}${selectedRunes[2]}0`));
    }
    
    const variant = currentSkill.variants[suffix];
    if(variant) {
        const descTxt = variant['desc_' + currentLang] || variant['desc_vi'] || "";
        document.getElementById('ui-desc').innerHTML = parseTags(descTxt);
    } else {
        document.getElementById('ui-desc').innerHTML = `(No data for ID: ${currentSkill.prefix}${suffix})`;
    }
}

// --- 4. CORE PARSER ---

function scanSkill(d) {
    if(d && typeof d === 'object'){
        let id = null;
        if (d.ID && d.ID.Value) id = d.ID.Value;
        
        if(id) {
            let entry = iconDb[parseInt(id)] || {};
            if(d.SkillIcon) {
                let iconPath = d.SkillIcon.replace(/\\/g, '/'); 
                entry.icon = iconPath.split('/').pop().toLowerCase() + ".png";
            }
            if(d.NeedCoolTime !== undefined) entry.cd = d.NeedCoolTime;
            if(d.NeedCostMp !== undefined) entry.mp = d.NeedCostMp;
            if(d.SkillLvGroupId) entry.gid = d.SkillLvGroupId;
            if(d.SkillType) entry.type = d.SkillType;

            iconDb[parseInt(id)] = entry;
        }
        Object.values(d).forEach(v => scanSkill(v));
    }
}

// --- HÀM MỚI: QUÉT DATA TỪ SKILLLV.JSON ---
function scanSkillLvData(d, db) {
    if(d && typeof d === 'object'){
        if(d.SkillLvGroupId && d.SkillLv) {
            const gid = d.SkillLvGroupId;
            const lv = d.SkillLv;
            if(!db[gid]) db[gid] = {};
            db[gid][lv] = {
                mp: d.NeedCostMp,
                cd: d.NeedCoolTime
            };
        }
        Object.values(d).forEach(v => scanSkillLvData(v, db));
    }
}

function scanJson(d, db) {
    if(d && typeof d === 'object'){
        let id = null;
        if (d.ID && d.ID.Value) id = d.ID.Value;
        else if (d.ID && typeof d.ID !== 'object') id = d.ID;

        if(id) {
            let e = {};
            if(d.Values) { e.v = d.Values; e.gid = d.AbnormalEffectLvGroupId; } 
            else if(d.EffectValueList) { 
                e.v = d.EffectValueList; 
                e.gid = d.SkillEffectLvGroupId; 
                if(d.AggroAbsolute !== undefined) e.aggro = d.AggroAbsolute;
            } 
            else if(d.TargetCountMax !== undefined) { e.v = d.TargetCountMax; }
            db[parseInt(id)] = e;
        }
        Object.values(d).forEach(v => scanJson(v, db));
    }
}
function scanJsonLv(d, db, isAbn) {
    if(d && d.Properties && d.Properties.Data) {
        d.Properties.Data.forEach(i => {
            let gid = isAbn ? i.AbnormalEffectLevelGroupId : i.SkillEffectLvGroupId;
            if(!gid && i.SkillLvGroupId) gid = i.SkillLvGroupId;

            let lv = isAbn ? i.AbnormalEffectLevel : i.SkillEffectLv;
            let v = isAbn ? i.Values : i.EffectValueList;
            if(gid && lv && v) { 
                if(!db[gid]) db[gid] = {}; 
                db[gid][lv] = { v: v, mp: i.CostMPoint, cd: i.CoolTime }; 
            }
        });
    }
}

const fmt = v => {
    if (isNaN(v)) return v; 
    return new Intl.NumberFormat('vi-VN').format(v); 
};

function getVals(db, dbLv, id, currentLv) {
    let entry = db[id];
    if(!entry) return null;
    let vals = entry.v;
    if(currentLv > 1 && entry.gid && dbLv[entry.gid] && dbLv[entry.gid][currentLv]) vals = dbLv[entry.gid][currentLv].v;
    if(Array.isArray(vals)) return vals.map(x => parseFloat(x) || x); 
    return vals;
}

// --- CẬP NHẬT: LẤY MP/CD TỪ SKILLLV.JSON ---
function getMeta(id, currentLv) {
    let mp = 0, cd = 0;
    let gid = null;

    const skillData = iconDb[id];
    if(skillData) {
        if(skillData.mp !== undefined) mp = skillData.mp;
        if(skillData.cd !== undefined) cd = skillData.cd;
        if(skillData.gid) gid = skillData.gid;
    }

    // ƯU TIÊN: Lấy từ dbSkillLv (File SkillLv.json)
    if(gid && dbSkillLv[gid] && dbSkillLv[gid][currentLv]) {
        const lvData = dbSkillLv[gid][currentLv];
        if(lvData.mp !== undefined) mp = lvData.mp;
        if(lvData.cd !== undefined) cd = lvData.cd;
    } 

    if(cd > 1000) cd /= 1000;
    return { mp, cd };
}

function parseTags(txt) {
    if(!txt) return "";
    txt = txt.replace(/<chat_combat>|<unique>|<copy_chat>|<\/chat_combat>|<\/unique>|<\/copy_chat>|<\/>/g, '').replace(/\{LF\}/g, '\n');

    return txt.replace(/\{(se_dmg|se_abe_dmg|se_abe|se_|abe|se|sef):([^}]+)\}/g, (match, tagType, content) => {
        const parts = content.split(':');
        const id = parseInt(parts[0]);

        if(tagType === 'sef') {
            if(dbFilter[id]) return `<span class="val">${dbFilter[id].v}</span>`;
            return "?";
        }

        if(tagType === 'se_abe_dmg') {
            const id2 = parseInt(parts[1]); 
            const type = parts[2];
            const vals = getVals(dbAbn, dbAbnLv, id2, currentLevel);
            if(!vals) return `<span class="bad">?</span>`;

            if(type.includes("Hot")) {
                const idx = type.includes("Min") ? 1 : 2; 
                return `<span class="heal">${fmt(vals[idx])}</span>`;
            } else {
                const isMin = type.includes("Min");
                const baseIdx = isMin ? 3 : 4;
                const scaleIdx = isMin ? 5 : 6;
                const base = parseFloat(vals[baseIdx]) / 10;
                const scale = parseFloat(vals[scaleIdx]) / 1000;
                let res = "";
                if(scale > 0) res += `${fmt(scale)}% ATK`; 
                if(base > 0) res += (res ? " + " : "") + fmt(base);
                res = `[${res || "0"}]`;
                return `<span class="val">${res || "0"}</span>`;
            }
        }
        
        if(tagType === 'se_abe') {
            const id2 = parseInt(parts[1]); 
            const valKey = parts[2]; 
            const op = parts[3]; 

            const vals = getVals(dbAbn, dbAbnLv, id2, currentLevel);
            if(!vals) return `<span class="bad">?</span>`;
            const idxMatch = valKey.match(/(\d+)$/);
            if(!idxMatch) return "?";
            const idx = parseInt(idxMatch[1], 10) - 1;
            if(vals[idx] === undefined) return "?";
            let val = parseFloat(vals[idx]);
            if(isNaN(val)) return `<span class="text">${vals[idx]}</span>`;
            if(op === 'divide100') val /= 100;
            return `<span class="val">${fmt(val)}</span>`; 
        }

        if(tagType === 'se_' || tagType === 'se_dmg') {
            const vals = getVals(dbDmg, dbDmgLv, id, currentLevel);
            if(!vals) return `<span class="bad">?</span>`;
            const type = parts[1]; 
            if(type.includes("Heal")) {
                const idx = type.includes("Min") ? 0 : 1;
                return `<span class="heal">${fmt(vals[idx])}</span>`;
            } else {
                const isMin = type.includes("Min");
                const baseIdx = isMin ? 0 : 1;
                const scaleIdx = isMin ? 2 : 3;
                const base = parseFloat(vals[baseIdx]);
                const scale = parseFloat(vals[scaleIdx]) / 100;
                let res = "";
                if(scale > 0) res += `${fmt(scale)}% ATK`; 
                if(base > 0) res += (res ? " + " : "") + fmt(base);
                res = `[${res || "0"}]`;
                return `<span class="val">${res}</span>`;
            }
        }

        if(tagType === 'abe' || tagType === 'se') {
            if (tagType === 'se' && parts[1] === 'aggro_absolute') {
                const entry = dbDmg[id];
                if (entry && entry.aggro !== undefined) return `<span class="val">${fmt(entry.aggro)}</span>`;
                return "?";
            }
            const db = (tagType === 'abe') ? dbAbn : dbDmg;
            const dbLv = (tagType === 'abe') ? dbAbnLv : dbDmgLv;
            const vals = getVals(db, dbLv, id, currentLevel);
            if(!vals) return `<span class="bad">?</span>`;
            const valKey = parts[1];
            const idxMatch = valKey.match(/(\d+)$/); 
            if(!idxMatch) return "?";
            const idx = parseInt(idxMatch[1], 10) - 1;
            if(vals[idx] === undefined) return "?";
            let val = parseFloat(vals[idx]);
            if(isNaN(val)) return `<span class="text">${vals[idx]}</span>`;
            const op = parts[2] || ""; 
            if(op === 'time') { val /= 1000; return `<span class="time">${fmt(val)}s</span>`; }
            else if(op === 'divide100') { val /= 100; return `<span class="val">${fmt(val)}</span>`; }
            else if(op === 'divide100abs') { val = Math.abs(val) / 100; return `<span class="val">${fmt(val)}</span>`; }
            return `<span class="val">${fmt(val)}</span>`;
        }
        return match;
    });
}

function getIcon(cls, fullId) {
    const idNum = parseInt(fullId.replace(/[^0-9]/g, ''));
    if(iconDb[idNum] && iconDb[idNum].icon) {
        // LOCAL PATH: ./icons/
        return `./icons/${iconDb[idNum].icon}`;
    }
    const cCode = CLASS_CODE[cls] || cls.substring(0,2).toLowerCase();
    const idx = fullId.substring(2,4); 
    const pad3 = idx.padStart(3,'0');
    const opts = [`icon_${cCode}_skill_${pad3}.png`, `icon_${cCode}_skill_${idx}.png`, `skill_${cCode}_${idx}.png`, `${fullId}.png`];
    for(let o of opts) if(iconMap[o]) return `./icons/${o}`;
    return null;
}

// --- TAB SWITCH LOGIC ---
function switchTab(mode) {
    const btnSearch = document.getElementById('btn-tab-search');
    const btnCompare = document.getElementById('btn-tab-compare');
    const viewSearch = document.getElementById('view-search');
    const viewCompare = document.getElementById('view-compare');

    if (mode === 'search') {
        // Active nút Tra Cứu
        btnSearch.classList.add('active');
        btnCompare.classList.remove('active');
        
        // Hiện giao diện Tra Cứu
        viewSearch.style.display = 'flex';
        viewCompare.style.display = 'none';
    } else {
        // Active nút So Sánh
        btnSearch.classList.remove('active');
        btnCompare.classList.add('active');
        
        // Ẩn giao diện Tra Cứu, Hiện So Sánh
        viewSearch.style.display = 'none';
        viewCompare.style.display = 'block';
        
        // --- [QUAN TRỌNG] DÒNG NÀY KÍCH HOẠT FILE COMPARE.JS ---
        // Nếu không có dòng này, dropdown sẽ trống trơn
        if(typeof initCompareTab === 'function') {
            initCompareTab(); 
        } else {
            console.error("Lỗi: Không tìm thấy hàm initCompareTab. Kiểm tra file compare.js!");
        }
    }
}