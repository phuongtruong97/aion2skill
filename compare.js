// --- COMPARE MODULE: FINAL (FIXED STIGMA LOGIC 4 RUNES) ---

let DATA_A = null;
let DATA_B = null;
let g_changes = []; 
let simA = null;
let simB = null;

// --- 1. KHỞI TẠO ---
async function initCompareTab() {
    try {
        const res = await fetch('./data/versions.json');
        if(!res.ok) return;
        const versions = await res.json();
        const selA = document.getElementById('ver-select-a');
        const selB = document.getElementById('ver-select-b');
        selA.innerHTML = ""; selB.innerHTML = "";
        versions.forEach(v => {
            selA.innerHTML += `<option value="${v.id}">${v.name}</option>`;
            selB.innerHTML += `<option value="${v.id}">${v.name}</option>`;
        });
        if(versions.length > 1) { selA.value = versions[1].id; selB.value = versions[0].id; }
    } catch(e) {}
}

async function runCompareProcess() {
    const idA = document.getElementById('ver-select-a').value;
    const idB = document.getElementById('ver-select-b').value;
    if(!idA || !idB) return alert("Vui lòng chọn 2 phiên bản!");

    // 1. Reset giao diện
    document.getElementById('compare-results-area').style.display = 'none';
    document.getElementById('diff-grouped-list').innerHTML = "";
    document.getElementById('compare-progress-container').style.display = 'block';
    
    const bar = document.getElementById('progress-bar');
    const txt = document.getElementById('progress-text');

    // 2. KHỞI TẠO HIỆU ỨNG GIẢ (FAKE LOADING)
    // Mẹo: Dùng setInterval để tự động tăng % trong khi chờ tải thật
    let percent = 5;
    bar.style.width = percent + '%';
    txt.innerText = "Đang tải dữ liệu...";

    const fakeTimer = setInterval(() => {
        // Mỗi lần nhảy ngẫu nhiên từ 5% đến 15%
        if (percent < 90) { 
            percent += Math.floor(Math.random() * 10) + 5; 
            if(percent > 90) percent = 90; // Chặn lại ở 90% nếu chưa tải xong
            bar.style.width = percent + '%';
        }
    }, 200); // Cứ 0.2 giây nhảy 1 lần

    try {
        // 3. TẢI DỮ LIỆU THẬT (Giữ nguyên Promise.all để ko bị lỗi logic)
        const [dA, dB] = await Promise.all([ loadAndProcessData(idA), loadAndProcessData(idB) ]);
        
        // 4. KHI TẢI XONG
        clearInterval(fakeTimer); // Dừng đồng hồ giả
        bar.style.width = '100%'; // Ép lên 100%
        txt.innerText = "Xử lý hoàn tất!";

        DATA_A = dA; DATA_B = dB;
        
        // setTimeout nhỏ để trình duyệt kịp vẽ thanh 100% trước khi hiện kết quả
        setTimeout(() => {
            scanChangesGrouped(DATA_A, DATA_B);
        }, 100);

    } catch(e) { 
        clearInterval(fakeTimer); // Nhớ tắt timer nếu lỗi
        console.error(e); 
        alert("Lỗi: " + e.message); 
        document.getElementById('compare-progress-container').style.display = 'none';
    }
}

// --- 2. LOAD DATA (NO CACHE) ---
async function loadAndProcessData(folderName) {
    const db = { iconDb: {}, dbDmg: {}, dbAbn: {}, dbSkillLv: {}, dbDmgLv: {}, dbAbnLv: {}, dbFilter: {}, skillMap: {} };
    // Dùng data_min (bản đã lọc field, nhẹ hơn ~86%) thay cho data gốc của game
    const path = `./data_min/${folderName}`;
    const ts = Date.now(); 

    const files = [
        { name: 'Skill.json', fn: scanSkill_C, target: db.iconDb },
        { name: 'SkillLv.json', fn: scanSkillLv_C, target: db.dbSkillLv },
        { name: 'SkillEffect.json', fn: scanJson_C, target: db.dbDmg },
        { name: 'SkillAbnormalEffect.json', fn: scanJson_C, target: db.dbAbn },
        { name: 'SkillEffectLv.json', fn: (d,t)=>scanJsonLv_C(d,t,false), target: db.dbDmgLv },
        { name: 'SkillAbnormalEffectLv.json', fn: (d,t)=>scanJsonLv_C(d,t,true), target: db.dbAbnLv },
        { name: 'SkillEffectFilter.json', fn: scanJson_C, target: db.dbFilter }
    ];
    
    await Promise.all(files.map(async f => {
        try {
            const res = await fetch(`${path}/${f.name}?t=${ts}`);
            if(res.ok) {
                const json = await res.json();
                f.fn(json, f.target);
            }
        } catch(e) { console.error(`Lỗi xử lý ${f.name}:`, e); }
    }));

    try {
        const res = await fetch(`${path}/text.xlsx?t=${ts}`);
        if(res.ok) {
            const buf = await res.arrayBuffer();
            const wb = XLSX.read(new Uint8Array(buf), {type:'array'});
            db.skillMap = processExcelSmart(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1}), db.iconDb);
        }
    } catch(e) { console.error('Lỗi xử lý text.xlsx:', e); }
    
    return db;
}

// --- 3. LOGIC SO SÁNH ---
// --- 3. LOGIC SO SÁNH (NÂNG CẤP: EXCEL TEXT + JSON VALUES) ---

// Helper: So sánh 2 mảng Values (JSON)
function isArrayDiff(arrA, arrB) {
    if (!arrA && !arrB) return false; // Cả 2 đều null -> không đổi
    if (!arrA || !arrB) return true;  // 1 có 1 không -> có đổi
    if (arrA.length !== arrB.length) return true; // Độ dài khác nhau -> có đổi
    // So sánh từng phần tử
    for (let i = 0; i < arrA.length; i++) {
        if (String(arrA[i]) !== String(arrB[i])) return true;
    }
    return false;
}

// Helper: Trích xuất ID từ chuỗi mô tả (Regex giống trong SkillSimulator)
function extractJsonIds(desc) {
    const ids = [];
    if (!desc) return ids;
    // Regex tìm các tag như {se_dmg:123}, {se_abe:456:Type}, {sef:789}
    const regex = /\{(se_dmg|se_abe_dmg|se_abe|se_|abe|se|sef):([^}]+)\}/g;
    let match;
    while ((match = regex.exec(desc)) !== null) {
        const tagType = match[1];
        const content = match[2];
        const parts = content.split(':');
        
        let id = parseInt(parts[0]); // Lấy ID đầu tiên
        let lookupType = 'dmg'; // Mặc định tìm trong SkillEffect (dbDmg)

        // Phân loại tìm trong dbAbn hay dbDmg
        if (tagType.includes('abe') || tagType.includes('sef') === false && content.includes('abe')) {
            lookupType = 'abn';
            // Trường hợp tag {se_abe_dmg:ID...} hoặc {se_abe:ID...}
             if (tagType === 'se_abe' && parts[1]) id = parseInt(parts[1]); // Fix cho trường hợp se_abe đặc biệt nếu có
        } else if (tagType === 'sef') {
            lookupType = 'filter';
        }

        // Chỉ push nếu ID hợp lệ
        if (!isNaN(id)) {
            ids.push({ id: id, type: lookupType });
        }
    }
    return ids;
}

// Helper: Kiểm tra thay đổi Value trong JSON
function checkJsonValueDiff(desc, dbA, dbB) {
    const refs = extractJsonIds(desc);
    if (refs.length === 0) return false;

    for (let item of refs) {
        let valA = null, valB = null;

        // Lấy dữ liệu từ DB A
        if (item.type === 'abn') valA = dbA.dbAbn[item.id] ? dbA.dbAbn[item.id].v : null;
        else if (item.type === 'filter') valA = dbA.dbFilter[item.id] ? dbA.dbFilter[item.id].v : null;
        else valA = dbA.dbDmg[item.id] ? dbA.dbDmg[item.id].v : null; // Default dbDmg

        // Lấy dữ liệu từ DB B
        if (item.type === 'abn') valB = dbB.dbAbn[item.id] ? dbB.dbAbn[item.id].v : null;
        else if (item.type === 'filter') valB = dbB.dbFilter[item.id] ? dbB.dbFilter[item.id].v : null;
        else valB = dbB.dbDmg[item.id] ? dbB.dbDmg[item.id].v : null;

        // So sánh
        if (isArrayDiff(valA, valB)) {
            // console.log(`Diff detected at ID ${item.id} (${item.type}):`, valA, "=>", valB); // Debug nếu cần
            return true;
        }
    }
    return false;
}

function scanChangesGrouped(dbA, dbB) {
    g_changes = [];
    const uiLang = document.getElementById('ui-lang') ? document.getElementById('ui-lang').value : 'vi';
    const keysB = Object.keys(dbB.skillMap);
    
    keysB.forEach(prefix => {
        const skillB = dbB.skillMap[prefix];
        const skillA = dbA.skillMap[prefix];
        
        let type = null;
        if(!skillA) {
            type = 'new';
        } else {
            let hasDiff = false;
            
            // 1. So sánh Tên (Excel)
            const nA = (uiLang === 'en' ? skillA.name_en : skillA.name_vi) || "";
            const nB = (uiLang === 'en' ? skillB.name_en : skillB.name_vi) || "";
            if (nA !== nB) hasDiff = true;

            // 2. So sánh Mô tả (Excel)
            const vA0 = skillA.variants["0000"] || {};
            const vB0 = skillB.variants["0000"] || {};
            const dA = (uiLang === 'en' ? vA0.desc_en : vA0.desc_vi) || "";
            const dB = (uiLang === 'en' ? vB0.desc_en : vB0.desc_vi) || "";
            if (dA !== dB) hasDiff = true;

            // 3. So sánh Rune Texts (Excel)
            const runeKeys = new Set([...Object.keys(skillA.runes), ...Object.keys(skillB.runes)]);
            runeKeys.forEach(rKey => {
                const rA = skillA.runes[rKey] || { en: "", vi: "" };
                const rB = skillB.runes[rKey] || { en: "", vi: "" };
                const txtA = (uiLang === 'en' ? rA.en : rA.vi) || "";
                const txtB = (uiLang === 'en' ? rB.en : rB.vi) || "";
                if (txtA !== txtB) hasDiff = true;
            });

            // 4. [NEW] So sánh Giá trị JSON (Hidden changes)
            // Quét các ID có trong mô tả của bản B, rồi so sánh giá trị đó trong A và B
            // Chúng ta dùng description tiếng Anh để quét vì tag thường đồng nhất
            const descToScan = (vB0.desc_en || vB0.desc_vi || ""); 
            if (!hasDiff) { // Nếu chưa thấy khác biệt ở Text, mới cần soi sâu vào JSON
                if (checkJsonValueDiff(descToScan, dbA, dbB)) {
                    hasDiff = true;
                }
            }
            
            // Quét thêm cả Rune Desc JSON nếu cần (Tuỳ chọn, tốn time hơn chút)
            if (!hasDiff) {
                 runeKeys.forEach(rKey => {
                    const rB = skillB.runes[rKey];
                    if (rB) {
                        const rDesc = rB.en || rB.vi || "";
                        if (checkJsonValueDiff(rDesc, dbA, dbB)) hasDiff = true;
                    }
                 });
            }

            if (hasDiff) type = 'change';
        }

        if(type) {
            g_changes.push({
                prefix: prefix,
                name: skillB[uiLang === 'en' ? 'name_en' : 'name_vi'],
                type: type,
                baseId: skillB.baseId,
                cls: skillB.className,
                skillDataB: skillB,
                skillDataA: skillA
            });
        }
    });

    document.getElementById('progress-bar').style.width = '100%';
    document.getElementById('progress-text').innerText = `Xong! Tìm thấy ${g_changes.length} thay đổi.`;
    setTimeout(() => {
        document.getElementById('compare-progress-container').style.display = 'none';
        renderResults();
    }, 300);
}

function renderResults() {
    const container = document.getElementById('diff-grouped-list');
    document.getElementById('compare-results-area').style.display = 'flex';
    container.innerHTML = "";

    if (g_changes.length === 0) {
        container.innerHTML = "<div style='padding:20px; text-align:center; color:#4caf50;'>Đôi khi tiếng Việt chưa kịp cập nhật. Hãy đổi sang tiếng Anh để kiểm tra thay đổi.</div>";
        return;
    }

    const grouped = {};
    g_changes.forEach(c => { if(!grouped[c.cls]) grouped[c.cls]=[]; grouped[c.cls].push(c); });

    Object.keys(grouped).sort().forEach(cls => {
        const div = document.createElement('div');
        div.className = 'class-group';
        div.innerHTML = `<div class="class-group-header">${cls} (${grouped[cls].length})</div>`;
        const grid = document.createElement('div');
        grid.className = 'skill-grid';
        grouped[cls].forEach(item => {
            const btn = document.createElement('div');
            btn.className = 'skill-btn';
            const iconUrl = getIcon_C(item.baseId, DATA_B) || getIcon_C(item.baseId, DATA_A);
            if(iconUrl) btn.style.backgroundImage = `url('${iconUrl}')`;
            btn.innerHTML = `<div class="diff-badge ${item.type==='new'?'new':''}">${item.type==='new'?'NEW':'MOD'}</div>
                             <div class="id-label">${item.prefix}</div>`;
            btn.onclick = () => openCompareModal(item);
            grid.appendChild(btn);
        });
        div.appendChild(grid);
        container.appendChild(div);
    });
}

function openCompareModal(item) {
    const modal = document.getElementById('compare-modal');
    modal.style.display = 'flex';
    
    const titleEl = document.getElementById('modal-skill-name');
    if(titleEl) titleEl.innerText = item.name;
    const idEl = document.getElementById('modal-skill-id');
    if(idEl) idEl.innerText = `ID: ${item.baseId}`;
    
    const iconUrl = getIcon_C(item.baseId, DATA_B) || getIcon_C(item.baseId, DATA_A);
    const iconEl = document.getElementById('modal-icon');
    if(iconEl) iconEl.style.backgroundImage = iconUrl ? `url('${iconUrl}')` : 'none';

    const targetA = document.getElementById('modal-content-a') || document.getElementById('card-container-a');
    const targetB = document.getElementById('modal-content-b') || document.getElementById('card-container-b');
    
    if(targetA) targetA.innerHTML = ""; 
    if(targetB) targetB.innerHTML = "";

    simA = null; simB = null;

    if (item.type !== 'new' && item.skillDataA) {
        simA = new SkillSimulator(targetA, item.skillDataA, DATA_A, "A");
    } else if(targetA) {
        targetA.innerHTML = "<div style='display:flex;height:100%;align-items:center;justify-content:center;color:#666'>(Không tồn tại)</div>";
    }
    
    if(targetB) {
        simB = new SkillSimulator(targetB, item.skillDataB, DATA_B, "B");
    }
    
    // Sync Level
    if(simA && simB) {
        simA.onLevelChange = (lv) => simB.setLevel(lv);
        simB.onLevelChange = (lv) => simA.setLevel(lv);
    }
}

function closeCompareModal() { document.getElementById('compare-modal').style.display = 'none'; }

// --- 5. CLASS SKILL SIMULATOR (LOGIC 4 RUNES = STIGMA) ---
class SkillSimulator {
    constructor(container, skillData, dbFull, uniqueId) {
        this.container = container;
        this.skill = skillData;
        this.db = dbFull;
        this.id = uniqueId;
        this.currentLevel = 1;
        this.selectedRunes = []; 
        this.lang = document.getElementById('ui-lang') ? document.getElementById('ui-lang').value : 'vi';

        // --- PHÂN LOẠI SKILL DỰA TRÊN SỐ LƯỢNG RUNE ---
        const runeCount = Object.keys(this.skill.runes).length;
        
        this.type = "active"; // Mặc định là Active

        if(this.skill.skillType === "ESkillType::Passive") {
            this.type = "passive";
        } 
        else if (
            (this.skill.skillType && this.skill.skillType.toLowerCase().includes("stigma")) || 
            runeCount === 4 // <--- QUY TẮC BẠN YÊU CẦU: 4 SPEC = STIGMA
        ) {
            this.type = "stigma";
        }
        
        // Khởi tạo
        if(this.type === 'stigma') {
            this.autoUpdateStigmaRunes();
        }
        
        this.initStructure();
        this.update();
    }

    autoUpdateStigmaRunes() {
        this.selectedRunes = [];
        // Stigma: 4 Rune tương ứng level 5, 10, 15, 20
        if(this.currentLevel >= 5) this.selectedRunes.push(1);
        if(this.currentLevel >= 10) this.selectedRunes.push(2);
        if(this.currentLevel >= 15) this.selectedRunes.push(3);
        if(this.currentLevel >= 20) this.selectedRunes.push(4);
    }

    setLevel(lv) {
        this.currentLevel = parseInt(lv);
        const slider = this.container.querySelector(`#slider-${this.id}`);
        if(slider) slider.value = this.currentLevel;
        this.container.querySelector(`#lv-txt-${this.id}`).innerText = this.currentLevel;
        
        if(this.type === 'stigma') {
            this.autoUpdateStigmaRunes();
        }
        this.update();
    }

    toggleRune(runeKey) {
        if(this.type === 'stigma') return; 
        
        const idx = parseInt(runeKey);
        if(this.selectedRunes.includes(idx)) {
            this.selectedRunes = this.selectedRunes.filter(i => i !== idx);
        } else {
            if(this.selectedRunes.length < 3) this.selectedRunes.push(idx);
        }
        this.selectedRunes.sort((a,b)=>a-b);
        this.update();
    }

    removeRuneAtSlot(slotIndex) {
        if(this.type === 'stigma') return; 
        if(slotIndex < this.selectedRunes.length) {
            this.selectedRunes.splice(slotIndex, 1);
            this.update();
        }
    }

    initStructure() {
        this.container.innerHTML = `
            <div class="compare-skill-container" style="display:flex; flex-direction:column; height:100%; position:relative; overflow-y:auto; padding-right:5px;">
                <div class="level-control" style="margin-top:0;">
                    <span style="font-size:12px;color:#888">LV:</span>
                    <input type="range" min="1" max="30" value="1" id="slider-${this.id}">
                    <span class="level-val" id="lv-txt-${this.id}">1</span>
                </div>
                <div class="skill-meta" style="margin-bottom:15px;">
                    <div class="meta-item"><span>💧</span> <span class="meta-val" id="mp-${this.id}">--</span></div>
                    <div class="meta-item"><span>⏳</span> <span class="meta-val" id="cd-${this.id}">--</span></div>
                    <div style="margin-left:auto; color:var(--accent); font-weight:bold; font-size:12px; text-transform:uppercase;">${this.type}</div>
                </div>
                <div class="desc-area" id="desc-${this.id}"></div>
                <div class="spec-container" id="spec-cont-${this.id}" style="display:none; border-top:1px solid #333; padding-top:10px;">
                    <div class="spec-header">
                        <span>SPECIALIZATION</span>
                        <span id="spec-count-${this.id}" style="color:#888">0/3</span>
                    </div>
                    <div class="rune-slots" id="slots-${this.id}"></div>
                    <div class="rune-pool" id="runes-${this.id}"></div>
                </div>
            </div>`;

        this.container.querySelector(`#slider-${this.id}`).oninput = (e) => {
            const val = e.target.value;
            this.setLevel(val);
            if(this.onLevelChange) this.onLevelChange(val); 
        };
    }

    update() {
        const baseIdNum = parseInt(this.skill.baseId.replace(/[^0-9]/g, ''));
        let mp = 0, cd = 0, gid = null;
        const skillMeta = this.db.iconDb[baseIdNum];
        if(skillMeta) {
            if(skillMeta.mp !== undefined) mp = skillMeta.mp;
            if(skillMeta.cd !== undefined) cd = skillMeta.cd;
            gid = skillMeta.gid;
        }
        if(gid && this.db.dbSkillLv && this.db.dbSkillLv[gid] && this.db.dbSkillLv[gid][this.currentLevel]) {
            const lvData = this.db.dbSkillLv[gid][this.currentLevel];
            if(lvData.mp !== undefined) mp = lvData.mp;
            if(lvData.cd !== undefined) cd = lvData.cd;
        }
        if(cd > 1000) cd /= 1000;
        const fmt = new Intl.NumberFormat('vi-VN');
        this.container.querySelector(`#mp-${this.id}`).innerText = mp > 0 ? fmt.format(mp) : "--";
        this.container.querySelector(`#cd-${this.id}`).innerText = cd > 0 ? cd + "s" : "--";

        // UI
        const specCont = this.container.querySelector(`#spec-cont-${this.id}`);
        const slotDiv = this.container.querySelector(`#slots-${this.id}`);
        const runeDiv = this.container.querySelector(`#runes-${this.id}`);
        const countTxt = this.container.querySelector(`#spec-count-${this.id}`);
        const runeKeys = Object.keys(this.skill.runes).map(Number).sort();

        if (this.type === 'passive' || runeKeys.length === 0) {
             specCont.style.display = 'none';
        } else {
             specCont.style.display = 'block';
             
             if (this.type === 'active') {
                 // Active: Hiện Slot
                 slotDiv.style.display = 'flex';
                 countTxt.innerText = `${this.selectedRunes.length}/3`;
                 slotDiv.innerHTML = "";
                 for(let i=0; i<3; i++) {
                     const slot = document.createElement('div');
                     slot.className = 'slot';
                     if (i < this.selectedRunes.length) {
                         slot.classList.add('filled');
                         const rid = this.selectedRunes[i];
                         const runeIndex = runeKeys.indexOf(rid); 
                         const num = (runeIndex % 5) + 1;
                         const imgName = `icon_item_usable_stigma_el_a_c_00${num}.png`;
                         slot.innerHTML = `<div class="rune-img" style="background-image: url('./icons/${imgName}')"></div><div class="slot-remove">✕</div>`;
                         slot.onclick = () => this.removeRuneAtSlot(i);
                     } else {
                         slot.innerHTML = `<div class="slot-remove">✕</div>`; 
                     }
                     slotDiv.appendChild(slot);
                 }
             } else {
                 // Stigma: Ẩn Slot
                 slotDiv.style.display = 'none';
                 countTxt.innerText = "(Auto by Level)";
             }

             runeDiv.innerHTML = "";
             runeKeys.forEach((rid, idx) => {
                 const row = document.createElement('div');
                 row.className = 'rune-row';
                 
                 // Highlight
                 if (this.selectedRunes.includes(rid)) row.classList.add('active');
                 
                 let imgName = "";
                 if (this.type === 'stigma') imgName = "icon_specialized_skill_stigma_common_001.png";
                 else imgName = `icon_item_usable_stigma_el_a_c_00${(idx % 5) + 1}.png`;
                 
                 const runeTxt = this.skill.runes[rid][this.lang] || this.skill.runes[rid]['vi'];
                 row.innerHTML = `<div class="rune-img" style="background-image: url('./icons/${imgName}')"></div>
                                  <div class="rune-text">${this.parseTags(runeTxt)}</div>`;
                 
                 // Chỉ Active mới cho click chọn
                 if(this.type !== 'stigma') {
                     row.onclick = () => this.toggleRune(rid);
                 } else {
                     row.style.cursor = 'default';
                 }
                 runeDiv.appendChild(row);
             });
        }

        // Desc
        let suffix = "0000";
        if(this.selectedRunes.length > 0) {
            if(this.type === 'stigma') {
                suffix = `00${this.selectedRunes.length}0`;
            } else {
                if(this.selectedRunes.length===1) suffix = `00${this.selectedRunes[0]}0`;
                else if(this.selectedRunes.length===2) suffix = `0${this.selectedRunes[0]}${this.selectedRunes[1]}0`;
                else suffix = `${this.selectedRunes[0]}${this.selectedRunes[1]}${this.selectedRunes[2]}0`;
            }
        }
        
        let variant = this.skill.variants[suffix];
        if (!variant) variant = this.skill.variants["0000"];

        const rawDesc = variant ? (variant[this.lang === 'en' ? 'desc_en' : 'desc_vi'] || "") : "(No Data)";
        this.container.querySelector(`#desc-${this.id}`).innerHTML = this.parseTags(rawDesc);
    }

    getVals(dbTarget, dbLvTarget, id) {
        let entry = dbTarget[id];
        if(!entry) return null;
        let vals = entry.v;
        if(this.currentLevel > 1 && entry.gid && dbLvTarget[entry.gid] && dbLvTarget[entry.gid][this.currentLevel]) {
            vals = dbLvTarget[entry.gid][this.currentLevel].v;
        }
        if(Array.isArray(vals)) return vals.map(x => parseFloat(x) || x);
        return vals;
    }

    parseTags(txt) {
        if(!txt) return "";
        txt = txt.replace(/<chat_combat>|<unique>|<copy_chat>|<\/chat_combat>|<\/unique>|<\/copy_chat>|<\/>/g, '').replace(/\{LF\}/g, '\n');
        const fmt = (v) => new Intl.NumberFormat('vi-VN').format(v);

        return txt.replace(/\{(se_dmg|se_abe_dmg|se_abe|se_|abe|se|sef):([^}]+)\}/g, (match, tagType, content) => {
            const parts = content.split(':');
            const id = parseInt(parts[0]);

            if(tagType === 'sef') {
                if(this.db.dbFilter[id]) return `<span class="val">${this.db.dbFilter[id].v}</span>`;
                return "?";
            }

            if(tagType === 'se_abe_dmg') {
                const id2 = parseInt(parts[1]); 
                const type = parts[2];
                const vals = this.getVals(this.db.dbAbn, this.db.dbAbnLv, id2); 
                if(!vals) return `<span class="bad">?</span>`;

                if(type && type.includes("Hot")) {
                    const idx = type.includes("Min") ? 1 : 2; 
                    return `<span class="heal">${fmt(vals[idx])}</span>`;
                } else {
                    const isMin = type && type.includes("Min");
                    const baseIdx = isMin ? 3 : 4;
                    const scaleIdx = isMin ? 5 : 6;
                    const base = parseFloat(vals[baseIdx]) / 10;
                    const scale = parseFloat(vals[scaleIdx]) / 1000;
                    let res = "";
                    if(scale > 0) res += `${fmt(scale)}% ATK`; 
                    if(base > 0) res += (res ? " + " : "") + fmt(base);
                    return `<span class="val">[${res || "0"}]</span>`;
                }
            }

            if(tagType === 'se_' || tagType === 'se_dmg') {
                const vals = this.getVals(this.db.dbDmg, this.db.dbDmgLv, id);
                if(!vals) return `<span class="bad">?</span>`;
                const type = parts[1]; 
                if(type && type.includes("Heal")) {
                    const idx = type.includes("Min") ? 0 : 1;
                    return `<span class="heal">${fmt(vals[idx])}</span>`;
                } else {
                    const isMin = type && type.includes("Min");
                    const baseIdx = isMin ? 0 : 1;
                    const scaleIdx = isMin ? 2 : 3;
                    const base = parseFloat(vals[baseIdx]);
                    const scale = parseFloat(vals[scaleIdx]) / 100;
                    let res = "";
                    if(scale > 0) res += `${fmt(scale)}% ATK`; 
                    if(base > 0) res += (res ? " + " : "") + fmt(base);
                    return `<span class="val">[${res || "0"}]</span>`;
                }
            }
            
            let activeDb = this.db.dbDmg; let activeDbLv = this.db.dbDmgLv; let lookupId = id;
            if(tagType.includes('abe')) { activeDb = this.db.dbAbn; activeDbLv = this.db.dbAbnLv; }
            if(tagType === 'se_abe') lookupId = parseInt(parts[1]); 

            if(tagType === 'se' && parts[1] === 'aggro_absolute') {
                const entry = this.db.dbDmg[id];
                return entry && entry.aggro ? `<span class="val">${fmt(entry.aggro)}</span>` : "?";
            }

            const vals = this.getVals(activeDb, activeDbLv, lookupId);
            if(!vals) return `<span class="bad">?</span>`;

            let valKey = parts[1];
            if (tagType === 'se_abe') valKey = parts[2];
            const idxMatch = valKey.match(/(\d+)$/);
            if(!idxMatch) return "?";
            const idx = parseInt(idxMatch[1], 10) - 1;
            
            if(vals[idx] === undefined) return "?";
            let val = parseFloat(vals[idx]);
            if(isNaN(val)) return `<span class="text">${vals[idx]}</span>`;
            
            const op = parts[parts.length-1]; 
            if(op === 'time') { val /= 1000; return `<span class="time">${fmt(val)}s</span>`; }
            else if(op === 'divide100') { val /= 100; return `<span class="val">${fmt(val)}</span>`; }
            else if(op === 'divide100abs') { val = Math.abs(val)/100; return `<span class="val">${fmt(val)}</span>`; }
            return `<span class="val">${fmt(val)}</span>`;
        });
    }
}

function processExcelSmart(rows, iconDb) {
    if (!rows || rows.length < 2) return {};

    const headers = rows[0];
    let idxId = 0, idxEn = 1, idxVi = 2; 
    if (Array.isArray(headers)) {
        headers.forEach((h, i) => {
            if (typeof h === 'string') {
                const lower = h.toLowerCase().trim();
                if (lower.includes('alias') || lower === 'id') idxId = i;
                else if (lower.includes('english') || lower === 'en') idxEn = i;
                else if (lower.includes('viet') || lower === 'vn' || lower === 'vi') idxVi = i;
            }
        });
    }

    let rawById = {};
    const idRegex = /(\d{8})/; 

    for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        let strId = r[idxId] ? String(r[idxId]) : "";
        let strEn = r[idxEn] ? String(r[idxEn]) : "";
        let strVi = r[idxVi] ? String(r[idxVi]) : ""; 

        let m = strId.match(idRegex);
        if (!m && strVi) m = strVi.match(idRegex);
        if (!m && strEn) m = strEn.match(idRegex);

        if (m) {
            const fid = m[1]; 
            let cls = "General"; 
            if(strId.includes("CLERIC")) cls = "CLERIC";
            else if(strId.includes("ASSASSIN")) cls = "ASSASSIN";
            else if(strId.includes("ARCHER")) cls = "ARCHER";
            else if(strId.includes("TEMPLAR")) cls = "TEMPLAR";
            else if(strId.includes("GLADIATOR")) cls = "GLADIATOR";
            else if(strId.includes("CHANTER")) cls = "CHANTER";
            else if(strId.includes("ELEMENTALIST")) cls = "ELEMENTALIST";
            else if(strId.includes("RANGER")) cls = "RANGER";
            else if(strId.includes("SORCERER")) cls = "SORCERER";

            if (!rawById[cls]) rawById[cls] = {};
            if (!rawById[cls][fid]) rawById[cls][fid] = { id: fid };
            
            let type = "desc";
            if (strEn.length < 50 && !strEn.includes("<") && !strEn.includes("{")) type = "name";
            if (strId.toLowerCase().includes("name")) type = "name";
            if (strId.toLowerCase().includes("specialized_skill_desc")) type = "rune";

            if (type === 'name') {
                rawById[cls][fid].name_en = strEn; rawById[cls][fid].name_vi = strVi;
            } else if (type === 'rune') {
                rawById[cls][fid].rune_en = strEn; rawById[cls][fid].rune_vi = strVi;
            } else {
                rawById[cls][fid].desc_en = strEn; rawById[cls][fid].desc_vi = strVi;
            }
        }
    }

    let finalMap = {}; 
    for (let cls in rawById) {
        for (let fid in rawById[cls]) {
            const prefix = fid.substring(0, 4);
            const suffix = fid.substring(4); 

            if (!finalMap[prefix]) {
                finalMap[prefix] = { prefix: prefix, baseId: fid, className: cls, variants: {}, runes: {}, skillType: null };
                const idNum = parseInt(prefix + "0000");
                if (iconDb[idNum] && iconDb[idNum].type) finalMap[prefix].skillType = iconDb[idNum].type;
            }
            const g = finalMap[prefix];
            const item = rawById[cls][fid];

            if(suffix === '0000') {
                 if(item.name_en) g.name_en = item.name_en;
                 if(item.name_vi) g.name_vi = item.name_vi;
                 g.baseId = fid;
            }
            g.variants[suffix] = { desc_en: item.desc_en || "", desc_vi: item.desc_vi || "" };

            if (item.rune_en || item.rune_vi) {
                const idx = parseInt(suffix[2]); 
                if (!isNaN(idx) && idx > 0) {
                    g.runes[idx] = { en: item.rune_en, vi: item.rune_vi };
                }
            }
        }
    }
    return finalMap;
}

// --- SCAN HELPERS ---
function scanSkill_C(d, db) { if(d && typeof d === 'object'){ if (d.ID && d.ID.Value) { let entry = db[d.ID.Value] || {}; if(d.SkillIcon) entry.icon = d.SkillIcon.replace(/\\/g, '/').split('/').pop().toLowerCase() + ".png"; if(d.NeedCoolTime !== undefined) entry.cd = d.NeedCoolTime; if(d.NeedCostMp !== undefined) entry.mp = d.NeedCostMp; if(d.SkillLvGroupId) entry.gid = d.SkillLvGroupId; if(d.SkillType) entry.type = d.SkillType; db[d.ID.Value] = entry; } Object.values(d).forEach(v => scanSkill_C(v, db)); } }
function scanSkillLv_C(d, db) { if(d && typeof d==='object'){ if(d.SkillLvGroupId && d.SkillLv) { if(!db[d.SkillLvGroupId]) db[d.SkillLvGroupId] = {}; db[d.SkillLvGroupId][d.SkillLv] = { mp: d.NeedCostMp, cd: d.NeedCoolTime }; } Object.values(d).forEach(v => scanSkillLv_C(v, db)); } }
// Thay thế hàm scanJson_C cũ bằng hàm này
function scanJson_C(d, db) { 
    if(d && typeof d === 'object'){ 
        // Lấy ID
        let id = d.ID && d.ID.Value ? d.ID.Value : (d.ID && typeof d.ID!=='object' ? d.ID : null); 
        
        if(id) { 
            let e = {}; 
            // Ưu tiên 1: Values (Mảng) -> SkillAbnormalEffect
            if(d.Values) { 
                e.v = d.Values; 
                e.gid = d.AbnormalEffectLvGroupId; 
            } 
            // Ưu tiên 2: EffectValueList -> SkillEffect
            else if(d.EffectValueList) { 
                e.v = d.EffectValueList; 
                e.gid = d.SkillEffectLvGroupId; 
                if(d.AggroAbsolute !== undefined) e.aggro = d.AggroAbsolute;
            } 
            // Ưu tiên 3: TargetCountMax (Số đơn) -> SkillEffectFilter (SEF)
            else if(d.TargetCountMax !== undefined) { 
                e.v = d.TargetCountMax; 
            }
            // Fallback: TargetCount
            else if(d.TargetCount !== undefined) { 
                e.v = d.TargetCount; 
            }

            // Chỉ lưu nếu có dữ liệu giá trị
            if(e.v !== undefined) {
                db[parseInt(id)] = e; 
            }
        } 
        Object.values(d).forEach(v => scanJson_C(v, db)); 
    } 
}function scanJsonLv_C(d, db, isAbn) { if(d && d.Properties && d.Properties.Data) { d.Properties.Data.forEach(i => { let gid = isAbn ? i.AbnormalEffectLevelGroupId : i.SkillEffectLvGroupId; if(!gid && i.SkillLvGroupId) gid = i.SkillLvGroupId; let lv = isAbn ? i.AbnormalEffectLevel : i.SkillEffectLv; let v = isAbn ? i.Values : i.EffectValueList; if(gid && lv && v) { if(!db[gid]) db[gid] = {}; db[gid][lv] = { v: v }; } }); } }
function getIcon_C(fullId, db) { if (!fullId) return null; const m = fullId.match(/(\d{8})/); if(m && db.iconDb && db.iconDb[parseInt(m[1])] && db.iconDb[parseInt(m[1])].icon) return `./icons/${db.iconDb[parseInt(m[1])].icon}`; return null; }