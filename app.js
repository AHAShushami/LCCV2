// Konfigurasi Google Sheets (Gantikan dengan URL Web App anda)
const GOOGLE_SHEET_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxXgg-1o0IDo3QyQs0LDWytGbkjblnpODSW16FLfTMjg9TuUnQwfs15DEYKVfGhsgur/exec";

// Struktur Data dan Faktor Pelepasan
const EMISSION_FACTORS = {
    // Tenaga
    elektrik: 2.3, // kg CO2e / kWh
    genset_diesel: 2.65, // kg CO2e / l
    
    // Air
    air: 10, // kg CO2e / m³
    
    // Sisa (Faktor per tan, dalam data pengguna masukkan dalam kg, jadi faktor / 1000)
    sisa_buangan: 12, // kg CO2e / tan (bersamaan 0.012 kg/kg)
    sisa_kitar_semula: 26, // kg CO2e / tan (bersamaan 0.026 kg/kg)
    
    // Kehijauan (Sequestration) (kg CO2e / tahun)
    hutan: 14, // per hektar
    landskap: 10, // per hektar
    badan_air: 12, // per hektar
    pokok: 25, // per pokok (purata)
    
    // Mobiliti
    petrol: 1.8, // kg CO2e / l
    diesel: 2.65, // kg CO2e / l
    cng: 86.21, // kg CO2e / Mbtus
    kenderaan_elektrik_kwh: 0 // Asingkan untuk dielakkan pelepasan berganda jika dicas di bangunan
};

const MAIN_SECTIONS = [
    { id: 'asas', title: 'Maklumat Asas', icon: 'ph-buildings' },
    { id: 'tenaga', title: 'Tenaga (Elektrik)', icon: 'ph-lightning' },
    { id: 'air', title: 'Air', icon: 'ph-drop' },
    { id: 'sisa', title: 'Sisa Pepejal', icon: 'ph-trash' },
    { id: 'hijau', title: 'Kehijauan', icon: 'ph-tree' },
    { id: 'mobiliti', title: 'Mobiliti (Pengangkutan)', icon: 'ph-car' },
    { id: 'rumusan', title: 'Rumusan Laporan', icon: 'ph-chart-pie-slice' }
];

const MONTHS = ['Januari', 'Februari', 'Mac', 'April', 'Mei', 'Jun', 'Julai', 'Ogos', 'September', 'Oktober', 'November', 'Disember'];

let formData = {
    asas: { 
        negeri: '', 
        daerah: '', 
        jenisFasiliti: '', 
        namaFasiliti: '', 
        pbt: '', 
        tahunAsas: 2025, 
        tahunProjek: 2026, 
        populasi: 50 
    },
    tahun: {
        '2025': {
            tenaga: { elektrik: Array(12).fill(0), genset: 0 },
            air: { penggunaan: Array(12).fill(0) },
            sisa: { buangan: Array(12).fill(0), kitar_semula: Array(12).fill(0) },
            hijau: { hutan: 0, landskap: 0, badan_air: 0, pokok: 0 },
            mobiliti: { petrol: 0, diesel: 0, cng: 0, ev_kwh: 0 }
        },
        '2026': {
            tenaga: { elektrik: Array(12).fill(0), genset: 0 },
            air: { penggunaan: Array(12).fill(0) },
            sisa: { buangan: Array(12).fill(0), kitar_semula: Array(12).fill(0) },
            hijau: { hutan: 0, landskap: 0, badan_air: 0, pokok: 0 },
            mobiliti: { petrol: 0, diesel: 0, cng: 0, ev_kwh: 0 }
        }
    }
};

let activeSection = 'asas';
let activeYear = '2025';
let chartInstance = null; // Menyimpan instance line chart

// Fungsi untuk menukar tahun aktif tab pendaftaran
window.setActiveYear = function(year) {
    activeYear = year;
    
    const btn2025 = document.getElementById('btn-year-2025');
    const btn2026 = document.getElementById('btn-year-2026');
    
    if(year === '2025') {
        btn2025.className = "relative z-10 px-5 py-2 text-sm font-semibold rounded-lg transition-all text-emerald-700 bg-white shadow-sm ring-1 ring-slate-900/5";
        btn2026.className = "relative z-10 px-5 py-2 text-sm font-medium rounded-lg transition-all text-slate-500 hover:text-slate-700 hover:bg-slate-200/50";
    } else {
        btn2026.className = "relative z-10 px-5 py-2 text-sm font-semibold rounded-lg transition-all text-emerald-700 bg-white shadow-sm ring-1 ring-slate-900/5";
        btn2025.className = "relative z-10 px-5 py-2 text-sm font-medium rounded-lg transition-all text-slate-500 hover:text-slate-700 hover:bg-slate-200/50";
    }
    
    // Perbaharui paparan
    if(activeSection !== 'asas' && activeSection !== 'rumusan') {
        renderSection(activeSection);
    }
};

window.copyFromBaseline = function(section) {
    if (confirm("Borang akan menyalin nilai dari Garis Asas (2025) ke tahun ini. Pasti?")) {
        const d2025_json = JSON.stringify(formData.tahun['2025'][section]);
        formData.tahun['2026'][section] = JSON.parse(d2025_json);
        saveToLocalStorage();
        renderSection(section);
    }
};

function saveToLocalStorage() {
    localStorage.setItem('lcc_formData', JSON.stringify(formData));
}

function loadFromLocalStorage() {
    const saved = localStorage.getItem('lcc_formData');
    if (saved) {
        try {
            const parsed = JSON.parse(saved);
            if (parsed.tahun && parsed.tahun['2025'] && parsed.tahun['2026']) {
                 formData = parsed;
            }
        } catch (e) {
            console.error("Gagal membaca LocalStorage", e);
        }
    }
}


// Helper: Format nombor
const formatNumber = (num, decimals = 2) => Number(num).toLocaleString('ms-MY', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });

// Helper: Kira jumlah bulanan
const sumArray = (arr) => arr.reduce((a, b) => Number(a) + Number(b), 0);

// Initialize App
document.addEventListener('DOMContentLoaded', () => {
    loadFromLocalStorage(); // <-- Muat data jika ada
    initNavigation();
    renderSection(activeSection);
    
    // Event listeners
    document.getElementById('btn-export').addEventListener('click', exportJSON);
    document.getElementById('import-file').addEventListener('change', importJSON);
    document.getElementById('btn-word').addEventListener('click', exportWord);
    document.getElementById('btn-google').addEventListener('click', sendToGoogleSheet);
});

function initNavigation() {
    const nav = document.getElementById('sidebar-nav');
    nav.innerHTML = '';
    
    MAIN_SECTIONS.forEach(sec => {
        const btn = document.createElement('div');
        btn.className = `tab-btn ${sec.id === activeSection ? 'active' : ''}`;
        btn.innerHTML = `<i class="ph ${sec.icon} text-xl mr-3"></i> ${sec.title}`;
        btn.onclick = () => {
            activeSection = sec.id;
            updateNavUI();
            renderSection(sec.id);
        };
        nav.appendChild(btn);
    });
}

function updateNavUI() {
    const buttons = document.querySelectorAll('.tab-btn');
    buttons.forEach((btn, index) => {
        if (MAIN_SECTIONS[index].id === activeSection) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
}

function renderSection(id) {
    const sec = MAIN_SECTIONS.find(s => s.id === id);
    document.getElementById('section-title').textContent = sec.title;
    
    const toggleContainer = document.getElementById('year-toggle-container');
    if(id === 'asas' || id === 'rumusan') {
        toggleContainer.classList.add('hidden');
        toggleContainer.classList.remove('inline-flex');
    } else {
        toggleContainer.classList.remove('hidden');
        toggleContainer.classList.add('inline-flex');
    }
    
    const content = document.getElementById('content-area');
    // Remove fade-in to trigger animation reflow
    content.classList.remove('fade-in');
    void content.offsetWidth; 
    content.classList.add('fade-in');
    
    content.innerHTML = '';
    
    switch(id) {
        case 'asas': renderAsas(content); break;
        case 'tenaga': renderTenaga(content); break;
        case 'air': renderAir(content); break;
        case 'sisa': renderSisa(content); break;
        case 'hijau': renderHijau(content); break;
        case 'mobiliti': renderMobiliti(content); break;
        case 'rumusan': renderRumusan(content); break;
    }
}

// ==== VIEW RENDERERS ====

function renderAsas(container) {
    document.getElementById('section-subtitle').textContent = "Lengkapkan butiran asas serta lokasi hierarki fasiliti kesihatan.";
    
    let html = `
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
                <label class="input-label">Negeri</label>
                <select class="input-field" onchange="updateData('asas', 'negeri', this.value)">
                    <option value="">-- Pilih Negeri --</option>
                    <option value="Johor" ${formData.asas.negeri === 'Johor' ? 'selected' : ''}>Johor</option>
                    <option value="Kedah" ${formData.asas.negeri === 'Kedah' ? 'selected' : ''}>Kedah</option>
                    <option value="Kelantan" ${formData.asas.negeri === 'Kelantan' ? 'selected' : ''}>Kelantan</option>
                    <option value="Melaka" ${formData.asas.negeri === 'Melaka' ? 'selected' : ''}>Melaka</option>
                    <option value="Negeri Sembilan" ${formData.asas.negeri === 'Negeri Sembilan' ? 'selected' : ''}>Negeri Sembilan</option>
                    <option value="Pahang" ${formData.asas.negeri === 'Pahang' ? 'selected' : ''}>Pahang</option>
                    <option value="Perak" ${formData.asas.negeri === 'Perak' ? 'selected' : ''}>Perak</option>
                    <option value="Perlis" ${formData.asas.negeri === 'Perlis' ? 'selected' : ''}>Perlis</option>
                    <option value="Pulau Pinang" ${formData.asas.negeri === 'Pulau Pinang' ? 'selected' : ''}>Pulau Pinang</option>
                    <option value="Sabah" ${formData.asas.negeri === 'Sabah' ? 'selected' : ''}>Sabah</option>
                    <option value="Sarawak" ${formData.asas.negeri === 'Sarawak' ? 'selected' : ''}>Sarawak</option>
                    <option value="Selangor" ${formData.asas.negeri === 'Selangor' ? 'selected' : ''}>Selangor</option>
                    <option value="Terengganu" ${formData.asas.negeri === 'Terengganu' ? 'selected' : ''}>Terengganu</option>
                    <option value="WP Kuala Lumpur" ${formData.asas.negeri === 'WP Kuala Lumpur' ? 'selected' : ''}>WP Kuala Lumpur</option>
                    <option value="WP Labuan" ${formData.asas.negeri === 'WP Labuan' ? 'selected' : ''}>WP Labuan</option>
                    <option value="WP Putrajaya" ${formData.asas.negeri === 'WP Putrajaya' ? 'selected' : ''}>WP Putrajaya</option>
                </select>
            </div>
            <div>
                <label class="input-label">Daerah</label>
                <input type="text" class="input-field" placeholder="Cth: Johor Bahru" value="${formData.asas.daerah}" onchange="updateData('asas', 'daerah', this.value)">
            </div>
            <div>
                <label class="input-label">Jenis Fasiliti</label>
                <select class="input-field" onchange="updateData('asas', 'jenisFasiliti', this.value)">
                    <option value="">-- Pilih Jenis Fasiliti --</option>
                    <option value="Klinik Kesihatan" ${formData.asas.jenisFasiliti === 'Klinik Kesihatan' ? 'selected' : ''}>Klinik Kesihatan (KK)</option>
                    <option value="Pejabat Kesihatan Daerah" ${formData.asas.jenisFasiliti === 'Pejabat Kesihatan Daerah' ? 'selected' : ''}>Pejabat Kesihatan Daerah (PKD)</option>
                    <option value="Hospital" ${formData.asas.jenisFasiliti === 'Hospital' ? 'selected' : ''}>Hospital</option>
                    <option value="Institusi Kesihatan Lain" ${formData.asas.jenisFasiliti === 'Institusi Kesihatan Lain' ? 'selected' : ''}>Institusi Kesihatan Pergigian / Lain-lain</option>
                </select>
            </div>
            <div>
                <label class="input-label">Nama Fasiliti / Organisasi</label>
                <input type="text" class="input-field" placeholder="Cth: Klinik Kesihatan Mahmoodiah" value="${formData.asas.namaFasiliti}" onchange="updateData('asas', 'namaFasiliti', this.value)">
            </div>
            <div>
                <label class="input-label">Pihak Berkuasa Tempatan (PBT) Kawasan</label>
                <input type="text" class="input-field" placeholder="Cth: MBJB" value="${formData.asas.pbt}" onchange="updateData('asas', 'pbt', this.value)">
            </div>
            <div class="grid grid-cols-2 gap-4">
                <div>
                    <label class="input-label">Tahun Asas</label>
                    <input type="number" class="input-field" value="${formData.asas.tahunAsas}" onchange="updateData('asas', 'tahunAsas', this.value)">
                </div>
                <div>
                    <label class="input-label">Tahun Projek</label>
                    <input type="number" class="input-field" value="${formData.asas.tahunProjek}" onchange="updateData('asas', 'tahunProjek', this.value)">
                </div>
            </div>
            <div>
                <label class="input-label">Populasi (Kakitangan / Penghuni)</label>
                <input type="number" class="input-field" value="${formData.asas.populasi}" onchange="updateData('asas', 'populasi', this.value)">
            </div>
        </div>
    `;
    container.innerHTML = html;
}

function renderTenaga(container) {
    document.getElementById('section-subtitle').textContent = "Pelepasan karbon kesan daripada penggunaan sumber elektrik grid dan bahan bakar.";
    
    const dataSection = formData.tahun[activeYear].tenaga;
    let totalKwh = sumArray(dataSection.elektrik);
    let carbonElektrik = (totalKwh * EMISSION_FACTORS.elektrik).toFixed(2);
    let carbonGenset = (dataSection.genset * EMISSION_FACTORS.genset_diesel).toFixed(2);
    
    let copyBtnHtml = activeYear === '2026' ? `<button onclick="copyFromBaseline('tenaga')" class="mb-4 text-xs font-semibold text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded hover:bg-emerald-100 transition inline-flex items-center gap-2 border border-emerald-200"><i class="ph ph-copy text-sm"></i> Salin Nilai 2025 (Garis Asas) Ke Sini</button>` : '';

    let html = `
        ${copyBtnHtml}
        <div class="formula-box">
            <h4 class="font-bold text-slate-700 flex items-center gap-2 mb-2"><i class="ph-fill ph-calculator"></i> Info Kiraan Karbon</h4>
            <div class="flex flex-col gap-2">
                <span class="formula-text">Pelepasan Elektrik = <span class="formula-highlight">∑ Penggunaan Bulanan (kWh)</span> × <span class="formula-highlight">${EMISSION_FACTORS.elektrik}</span> kg CO₂e/kWh</span>
                <span class="formula-text">Pelepasan Genset = <span class="formula-highlight">Penggunaan Litre</span> × <span class="formula-highlight">${EMISSION_FACTORS.genset_diesel}</span> kg CO₂e/l</span>
            </div>
        </div>
        
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-8">
            <div class="col-span-2">
                <h3 class="font-semibold mb-4 border-b pb-2">Penggunaan Elektrik Bulanan (kWh) - ${activeYear}</h3>
                <div class="grid grid-cols-2 md:grid-cols-3 gap-4">
                    ${MONTHS.map((m, i) => `
                        <div>
                            <label class="text-xs text-slate-500 mb-1 block">${m}</label>
                            <input type="number" class="input-field py-1.5" value="${dataSection.elektrik[i]}" oninput="updateArrayData('tenaga', 'elektrik', ${i}, this.value); updateRealtimeCarbon('tenaga')">
                        </div>
                    `).join('')}
                </div>
                <div class="mt-6 p-4 bg-slate-50 rounded-lg flex justify-between items-center text-sm font-semibold text-slate-700">
                    <span>Jumlah Setahun: <span id="tenaga-total-val" class="text-primary">${formatNumber(totalKwh)}</span> kWh</span>
                    <span>Jejak Karbon: <span id="tenaga-carbon-val" class="text-red-500 text-lg">${formatNumber(carbonElektrik)}</span> kg CO₂e</span>
                </div>
            </div>
            <div>
                <h3 class="font-semibold mb-4 border-b pb-2">Set Penjana (Genset) - ${activeYear}</h3>
                <div>
                    <label class="input-label">Minyak Diesel Setahun (Liter)</label>
                    <input type="number" class="input-field" value="${dataSection.genset}" oninput="updateData('tenaga', 'genset', this.value); updateRealtimeCarbon('tenaga')">
                </div>
                <div class="mt-6 p-4 bg-slate-50 rounded-lg text-sm font-semibold text-slate-700">
                    <span>Jejak Karbon: <br><span id="genset-carbon-val" class="text-red-500 text-lg mt-1 inline-block">${formatNumber(carbonGenset)}</span> kg CO₂e</span>
                </div>
            </div>
        </div>
    `;
    container.innerHTML = html;
}

function renderAir(container) {
    document.getElementById('section-subtitle').textContent = "Pelepasan karbon melalui rantaian bekalan rawatan air bersih.";
    
    const dataSection = formData.tahun[activeYear].air;
    let totalAir = sumArray(dataSection.penggunaan);
    let carbonAir = (totalAir * EMISSION_FACTORS.air).toFixed(2);
    
    let copyBtnHtml = activeYear === '2026' ? `<button onclick="copyFromBaseline('air')" class="mb-4 text-xs font-semibold text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded hover:bg-emerald-100 transition inline-flex items-center gap-2 border border-emerald-200"><i class="ph ph-copy text-sm"></i> Salin Nilai 2025 (Garis Asas) Ke Sini</button>` : '';

    let html = `
        ${copyBtnHtml}
        <div class="formula-box">
            <h4 class="font-bold text-slate-700 flex items-center gap-2 mb-2"><i class="ph-fill ph-calculator"></i> Info Kiraan Karbon</h4>
            <div class="flex flex-col gap-2">
                <span class="formula-text">Pelepasan Bekalan Air = <span class="formula-highlight">∑ Penggunaan (m³)</span> × <span class="formula-highlight">${EMISSION_FACTORS.air}</span> kg CO₂e/m³</span>
            </div>
        </div>
        
        <div class="max-w-3xl">
            <h3 class="font-semibold mb-4 border-b pb-2">Penggunaan Air Bulanan (M³ Bil Domestik) - ${activeYear}</h3>
            <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
                ${MONTHS.map((m, i) => `
                    <div>
                        <label class="text-xs text-slate-500 mb-1 block">${m}</label>
                        <input type="number" class="input-field py-1.5" value="${dataSection.penggunaan[i]}" oninput="updateArrayData('air', 'penggunaan', ${i}, this.value); updateRealtimeCarbon('air')">
                    </div>
                `).join('')}
            </div>
            <div class="mt-8 p-5 bg-emerald-50 border border-emerald-100 rounded-xl flex justify-between items-center text-sm font-semibold text-slate-700">
                <span>Jumlah Penggunaan: <span id="air-total-val" class="text-primary text-lg">${formatNumber(totalAir)}</span> m³</span>
                <span class="text-right">Jejak Karbon: <br><span id="air-carbon-val" class="text-red-500 text-xl font-bold">${formatNumber(carbonAir)}</span> kg CO₂e</span>
            </div>
        </div>
    `;
    container.innerHTML = html;
}

function renderSisa(container) {
    document.getElementById('section-subtitle').textContent = "Pelepasan sisa solid ke tapak pelupusan berbanding hasil kitar semula.";
    
    const dataSection = formData.tahun[activeYear].sisa;
    let totalTapak = sumArray(dataSection.buangan);
    let totalKitar = sumArray(dataSection.kitar_semula);
    // Faktor formula sisa adalah per Tan. Berat diinput dalam KG.
    let carbonTapak = (totalTapak / 1000 * EMISSION_FACTORS.sisa_buangan).toFixed(2);
    let carbonKitar = (totalKitar / 1000 * EMISSION_FACTORS.sisa_kitar_semula).toFixed(2);
    
    let copyBtnHtml = activeYear === '2026' ? `<button onclick="copyFromBaseline('sisa')" class="mb-4 text-xs font-semibold text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded hover:bg-emerald-100 transition inline-flex items-center gap-2 border border-emerald-200"><i class="ph ph-copy text-sm"></i> Salin Nilai 2025 (Garis Asas) Ke Sini</button>` : '';

    let html = `
        ${copyBtnHtml}
        <div class="formula-box">
            <h4 class="font-bold text-slate-700 flex items-center gap-2 mb-2"><i class="ph-fill ph-calculator"></i> Info Kiraan Karbon</h4>
            <div class="flex flex-col gap-2">
                <span class="formula-text">Pelepasan Sisa Domestik = <span class="formula-highlight">∑ Sisa Buangan (kg) ÷ 1000</span> × <span class="formula-highlight">${EMISSION_FACTORS.sisa_buangan}</span> kg CO₂e/tan</span>
                <span class="formula-text text-emerald-700">Pelepasan Sisa Kitar Semula *(Estimate)* = <span class="formula-highlight">∑ Kitar Semula (kg) ÷ 1000</span> × <span class="formula-highlight">${EMISSION_FACTORS.sisa_kitar_semula}</span> kg CO₂e/tan</span>
            </div>
        </div>
        
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-8">
            <!-- Buangan ke Tapak Pelupusan -->
            <div>
                <h3 class="font-semibold mb-4 border-b border-rose-200 pb-2 text-rose-700">Sisa Pepejal ke Tapak Pelupusan (KG) - ${activeYear}</h3>
                <div class="grid grid-cols-2 lg:grid-cols-3 gap-3">
                    ${MONTHS.map((m, i) => `
                        <div>
                            <label class="text-[11px] text-slate-500 mb-1 block">${m}</label>
                            <input type="number" class="input-field py-1" value="${dataSection.buangan[i]}" oninput="updateArrayData('sisa', 'buangan', ${i}, this.value); updateRealtimeCarbon('sisa')">
                        </div>
                    `).join('')}
                </div>
                <div class="mt-4 p-3 bg-rose-50 rounded-lg text-sm text-center font-medium">
                    Jejak Karbon (Buangan): <span id="sisa-buangan-carbon" class="text-rose-600 font-bold">${formatNumber(carbonTapak)}</span> kg CO₂e
                </div>
            </div>
            
            <!-- Kitar Semula -->
            <div>
                <h3 class="font-semibold mb-4 border-b border-emerald-200 pb-2 text-emerald-700">Sisa Kitar Semula Terasing (KG) - ${activeYear}</h3>
                <div class="grid grid-cols-2 lg:grid-cols-3 gap-3">
                    ${MONTHS.map((m, i) => `
                        <div>
                            <label class="text-[11px] text-slate-500 mb-1 block">${m}</label>
                            <input type="number" class="input-field py-1" value="${dataSection.kitar_semula[i]}" oninput="updateArrayData('sisa', 'kitar_semula', ${i}, this.value); updateRealtimeCarbon('sisa')">
                        </div>
                    `).join('')}
                </div>
                <div class="mt-4 p-3 bg-emerald-50 rounded-lg text-sm text-center font-medium">
                    Jejak Karbon (Kitar Semula): <span id="sisa-kitar-carbon" class="text-emerald-600 font-bold">${formatNumber(carbonKitar)}</span> kg CO₂e
                </div>
            </div>
        </div>
    `;
    container.innerHTML = html;
}

function renderHijau(container) {
    document.getElementById('section-subtitle').textContent = "Sink Karbon Semulajadi: Pokok dan liputan hijau yang menyerap sisa karbon kawasan.";
    
    const dataSection = formData.tahun[activeYear].hijau;
    let serapan = (dataSection.hutan * EMISSION_FACTORS.hutan) +
                  (dataSection.landskap * EMISSION_FACTORS.landskap) +
                  (dataSection.badan_air * EMISSION_FACTORS.badan_air) +
                  (dataSection.pokok * EMISSION_FACTORS.pokok);
                  
    let copyBtnHtml = activeYear === '2026' ? `<button onclick="copyFromBaseline('hijau')" class="mb-4 text-xs font-semibold text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded hover:bg-emerald-100 transition inline-flex items-center gap-2 border border-emerald-200"><i class="ph ph-copy text-sm"></i> Salin Nilai 2025 (Garis Asas) Ke Sini</button>` : '';

    let html = `
        ${copyBtnHtml}
        <div class="formula-box !border-emerald-500 !bg-emerald-50">
            <h4 class="font-bold text-emerald-800 flex items-center gap-2 mb-2"><i class="ph-fill ph-tree"></i> Info Serapan (Sequestration) Positif</h4>
            <div class="flex flex-col gap-2">
                <span class="formula-text text-emerald-700">Serapan Hutan = <span class="formula-highlight">${EMISSION_FACTORS.hutan}</span> kg CO₂e/ha</span>
                <span class="formula-text text-emerald-700">Serapan Landskap = <span class="formula-highlight">${EMISSION_FACTORS.landskap}</span> kg CO₂e/ha</span>
                <span class="formula-text text-emerald-700">Serapan Badan Air = <span class="formula-highlight">${EMISSION_FACTORS.badan_air}</span> kg CO₂e/ha</span>
                <span class="formula-text text-emerald-700">Serapan Pokok (>5 thn) = <span class="formula-highlight">${EMISSION_FACTORS.pokok}</span> kg CO₂e/pokok</span>
            </div>
        </div>
        
        <div class="grid grid-cols-1 md:grid-cols-2 gap-8 max-w-4xl">
            <div class="space-y-4">
                <div>
                    <label class="input-label">Keluasan Hutan (Hektar / ha)</label>
                    <input type="number" class="input-field" value="${dataSection.hutan}" oninput="updateData('hijau', 'hutan', this.value); updateRealtimeCarbon('hijau')">
                </div>
                <div>
                    <label class="input-label">Keluasan Landskap / Rumput (Hektar / ha)</label>
                    <input type="number" class="input-field" value="${dataSection.landskap}" oninput="updateData('hijau', 'landskap', this.value); updateRealtimeCarbon('hijau')">
                </div>
                <div>
                    <label class="input-label">Keluasan Badan Air (Hektar / ha)</label>
                    <input type="number" class="input-field" value="${dataSection.badan_air}" oninput="updateData('hijau', 'badan_air', this.value); updateRealtimeCarbon('hijau')">
                </div>
                <div>
                    <label class="input-label">Bilangan Pokok Ditanam (> 5 tahun)</label>
                    <input type="number" class="input-field" value="${dataSection.pokok}" oninput="updateData('hijau', 'pokok', this.value); updateRealtimeCarbon('hijau')">
                </div>
            </div>
            
            <div class="glass-card rounded-xl p-8 flex flex-col items-center justify-center text-center bg-gradient-to-b from-white to-emerald-50">
                <i class="ph-fill ph-plant text-6xl text-emerald-400 mb-4 opacity-75"></i>
                <h3 class="text-slate-500 font-semibold mb-2 uppercase tracking-wide text-sm">Jumlah Karbon Diserap (${activeYear})</h3>
                <span id="hijau-serapan" class="text-4xl text-emerald-600 font-bold mt-2">
                    - ${formatNumber(serapan)} <span class="text-lg text-emerald-500">kg CO₂e</span>
                </span>
                <p class="text-xs text-slate-400 mt-4 leading-relaxed">Penyumbang karbon positif yang membersihkan udara persekitaran PBT/Institut.</p>
            </div>
        </div>
    `;
    container.innerHTML = html;
}

function renderMobiliti(container) {
    document.getElementById('section-subtitle').textContent = "Pelepasan bahan bakar pengangkutan syarikat/institusi (Skop 1).";
    
    const dataSection = formData.tahun[activeYear].mobiliti;
    let carbonMobiId = (dataSection.petrol * EMISSION_FACTORS.petrol) +
                       (dataSection.diesel * EMISSION_FACTORS.diesel) +
                       (dataSection.cng * EMISSION_FACTORS.cng);

    let copyBtnHtml = activeYear === '2026' ? `<button onclick="copyFromBaseline('mobiliti')" class="mb-4 text-xs font-semibold text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded hover:bg-emerald-100 transition inline-flex items-center gap-2 border border-emerald-200"><i class="ph ph-copy text-sm"></i> Salin Nilai 2025 (Garis Asas) Ke Sini</button>` : '';

    let html = `
        ${copyBtnHtml}
        <div class="formula-box">
            <h4 class="font-bold text-slate-700 flex items-center gap-2 mb-2"><i class="ph-fill ph-calculator"></i> Info Kiraan Karbon (Pengangkutan Bergerak)</h4>
            <div class="flex flex-col gap-2">
                <span class="formula-text">Pelepasan Petrol = <span class="formula-highlight">${EMISSION_FACTORS.petrol}</span> kg CO₂e/liter</span>
                <span class="formula-text">Pelepasan Diesel = <span class="formula-highlight">${EMISSION_FACTORS.diesel}</span> kg CO₂e/liter</span>
                <span class="formula-text">Pelepasan CNG = <span class="formula-highlight">${EMISSION_FACTORS.cng}</span> kg CO₂e/Mbtus</span>
            </div>
        </div>
        
        <div class="grid grid-cols-1 md:grid-cols-2 gap-x-12 gap-y-6 max-w-4xl">
            <div>
                <label class="input-label">Penggunaan Petrol Setahun (Liter) - ${activeYear}</label>
                <input type="number" class="input-field" value="${dataSection.petrol}" oninput="updateData('mobiliti', 'petrol', this.value); updateRealtimeCarbon('mobiliti')">
            </div>
            <div>
                <label class="input-label">Penggunaan Diesel Setahun (Liter) - ${activeYear}</label>
                <input type="number" class="input-field" value="${dataSection.diesel}" oninput="updateData('mobiliti', 'diesel', this.value); updateRealtimeCarbon('mobiliti')">
            </div>
            <div>
                <label class="input-label">Gas Asli Kenderaan (CNG) (Mbtus) - ${activeYear}</label>
                <input type="number" class="input-field" value="${dataSection.cng}" oninput="updateData('mobiliti', 'cng', this.value); updateRealtimeCarbon('mobiliti')">
            </div>
            <div class="col-span-1 border-t lg:border-t-0 p-4 lg:p-0 bg-slate-50 lg:bg-transparent rounded-lg">
                <h3 class="font-semibold text-slate-700 border-b pb-2 mb-4">Pengecasan EV (Estimasi Sifar untuk Skop 1)</h3>
                <label class="input-label">Penggunaan Pengecas (Charger) (kWh)</label>
                <input type="number" class="input-field bg-white" value="${dataSection.ev_kwh}" oninput="updateData('mobiliti', 'ev_kwh', this.value)">
                <span class="text-xs text-slate-400 mt-2 block">Nota: Penggunaan elektrik charger EV dikira di dalam bil keseluruhan institusi (Tenaga Tab) bagi mengelak pengiraan pelepasan berganda.</span>
            </div>
        </div>
        
        <div class="mt-8 pt-6 border-t font-semibold text-slate-700 flex items-center justify-between max-w-4xl">
            <span>Jumlah Pelepasan Skop 1 (Pengangkutan):</span>
            <span id="mobiliti-carbon-val" class="text-2xl text-red-500">${formatNumber(carbonMobiId)} <span class="text-sm">kg CO₂e</span></span>
        </div>
    `;
    container.innerHTML = html;
}

function calculateEmissions(year) {
    const data = formData.tahun[year];
    let t_energy = (sumArray(data.tenaga.elektrik) * EMISSION_FACTORS.elektrik) + (data.tenaga.genset * EMISSION_FACTORS.genset_diesel);
    let t_air = sumArray(data.air.penggunaan) * EMISSION_FACTORS.air;
    let t_waste = ((sumArray(data.sisa.buangan) / 1000) * EMISSION_FACTORS.sisa_buangan) + ((sumArray(data.sisa.kitar_semula) / 1000) * EMISSION_FACTORS.sisa_kitar_semula);
    let t_mobility = (data.mobiliti.petrol * EMISSION_FACTORS.petrol) + (data.mobiliti.diesel * EMISSION_FACTORS.diesel) + (data.mobiliti.cng * EMISSION_FACTORS.cng);
    
    let gross = t_energy + t_air + t_waste + t_mobility;
    let sink = (data.hijau.hutan * EMISSION_FACTORS.hutan) + (data.hijau.landskap * EMISSION_FACTORS.landskap) + (data.hijau.badan_air * EMISSION_FACTORS.badan_air) + (data.hijau.pokok * EMISSION_FACTORS.pokok);
    let net = gross - sink;
    
    return { t_energy, t_air, t_waste, t_mobility, gross, sink, net };
}

function calculateVariance(v2025, v2026) {
    if (v2025 === 0 && v2026 === 0) return { diff: 0, pct: 0, isReduction: true };
    if (v2025 === 0) return { diff: v2026, pct: 100, isReduction: false };
    let diff = v2026 - v2025;
    let pct = (diff / v2025) * 100;
    return { diff: Math.abs(diff), pct: Math.abs(pct), isReduction: diff <= 0 };
}

function renderVarianceBadge(varianceInfo) {
    if (varianceInfo.isReduction && (varianceInfo.diff > 0 || varianceInfo.pct === 0)) {
        return `<span class="bg-emerald-100 text-emerald-700 font-bold px-2 py-0.5 rounded text-[10px] ml-2 inline-block">📉 -${formatNumber(varianceInfo.pct, 1)}%</span>`;
    } else {
        return `<span class="bg-rose-100 text-rose-700 font-bold px-2 py-0.5 rounded text-[10px] ml-2 inline-block">📈 +${formatNumber(varianceInfo.pct, 1)}%</span>`;
    }
}

function renderRumusan(container) {
    document.getElementById('section-subtitle').textContent = "Pemandangan komprehensif perbandingan pelepasan gas rumah hijau kawasan anda berbanding garis asas (2025).";
    
    let d2025 = calculateEmissions('2025');
    let d2026 = calculateEmissions('2026');
    
    let varNet = calculateVariance(d2025.net, d2026.net);

    let html = `
        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
            <div class="glass-card p-6 rounded-xl text-center border-slate-200 shadow-sm relative overflow-hidden group">
                <div class="absolute right-0 top-0 p-4 opacity-10 blur-sm pointer-events-none transition-all group-hover:blur-none group-hover:opacity-20"><i class="ph-fill ph-target text-6xl"></i></div>
                <h4 class="text-slate-500 text-sm font-semibold uppercase tracking-wide">Pelepasan Asas (2025)</h4>
                <div class="text-3xl font-bold text-slate-800 mt-2">${formatNumber(d2025.net)}</div>
                <div class="text-xs text-slate-400">kg CO₂e</div>
            </div>
            
            <div class="glass-card p-6 rounded-xl text-center border-slate-200 shadow-sm relative overflow-hidden group">
                <div class="absolute right-0 top-0 p-4 opacity-10 blur-sm pointer-events-none transition-all group-hover:blur-none group-hover:opacity-20"><i class="ph-fill ph-flag-checkered text-6xl"></i></div>
                <h4 class="text-slate-500 text-sm font-semibold uppercase tracking-wide">Pencapaian Semasa (2026)</h4>
                <div class="text-3xl font-bold text-slate-800 mt-2">${formatNumber(d2026.net)}</div>
                <div class="text-xs text-slate-400">kg CO₂e</div>
            </div>

            <div class="glass-card p-6 rounded-xl text-center ${varNet.isReduction ? 'border-emerald-300 bg-emerald-50' : 'border-rose-300 bg-rose-50'} shadow-md">
                <i class="ph-fill ${varNet.isReduction ? 'ph-trend-down text-emerald-500' : 'ph-trend-up text-rose-500'} text-4xl mb-2"></i>
                <h4 class="text-slate-600 text-sm font-semibold uppercase tracking-wide">Perbezaan (Variance)</h4>
                <div class="text-4xl font-extrabold ${varNet.isReduction ? 'text-emerald-700' : 'text-rose-700'} mt-2">
                    ${varNet.isReduction && varNet.diff > 0 ? '-' : (varNet.diff === 0 ? '' : '+')}${formatNumber(varNet.pct, 1)}%
                </div>
                <div class="text-xs font-medium mt-1 ${varNet.isReduction ? 'text-emerald-600' : 'text-rose-600'}">
                    ${varNet.isReduction ? (varNet.diff === 0 ? 'Tiada Perubahan' : 'Penurunan Berjaya') : 'Peningkatan Gagal'}
                </div>
            </div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
            <div>
                <h3 class="font-bold text-slate-800 text-lg mb-4 border-b pb-2">Prestasi Sektor (kg CO₂e Kasar)</h3>
                <table class="data-table w-full rounded-lg overflow-hidden border border-slate-200 shadow-sm">
                    <thead>
                        <tr>
                            <th>Sektor</th>
                            <th class="text-right">Asas 2025</th>
                            <th class="text-right">Semasa 2026</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td class="font-medium"><i class="ph-fill ph-lightning text-yellow-500 mr-2"></i>Tenaga</td>
                            <td class="text-right text-slate-500">${formatNumber(d2025.t_energy)}</td>
                            <td class="text-right text-slate-700 font-bold">${formatNumber(d2026.t_energy)}<br>${renderVarianceBadge(calculateVariance(d2025.t_energy, d2026.t_energy))}</td>
                        </tr>
                        <tr>
                            <td class="font-medium"><i class="ph-fill ph-drop text-blue-500 mr-2"></i>Air</td>
                            <td class="text-right text-slate-500">${formatNumber(d2025.t_air)}</td>
                            <td class="text-right text-slate-700 font-bold">${formatNumber(d2026.t_air)}<br>${renderVarianceBadge(calculateVariance(d2025.t_air, d2026.t_air))}</td>
                        </tr>
                        <tr>
                            <td class="font-medium"><i class="ph-fill ph-trash text-orange-500 mr-2"></i>Sisa Pepejal</td>
                            <td class="text-right text-slate-500">${formatNumber(d2025.t_waste)}</td>
                            <td class="text-right text-slate-700 font-bold">${formatNumber(d2026.t_waste)}<br>${renderVarianceBadge(calculateVariance(d2025.t_waste, d2026.t_waste))}</td>
                        </tr>
                        <tr>
                            <td class="font-medium"><i class="ph-fill ph-car text-slate-500 mr-2"></i>Mobiliti</td>
                            <td class="text-right text-slate-500">${formatNumber(d2025.t_mobility)}</td>
                            <td class="text-right text-slate-700 font-bold">${formatNumber(d2026.t_mobility)}<br>${renderVarianceBadge(calculateVariance(d2025.t_mobility, d2026.t_mobility))}</td>
                        </tr>
                        <tr class="bg-emerald-50/50">
                            <td class="font-bold text-emerald-800"><i class="ph-fill ph-tree text-emerald-600 mr-2"></i>- Serapan Sink</td>
                            <td class="text-right text-slate-500 text-sm">(${formatNumber(d2025.sink)})</td>
                            <td class="text-right text-emerald-700 font-bold text-sm">(${formatNumber(d2026.sink)})</td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <div class="glass-card p-5 rounded-xl border border-slate-200 flex flex-col">
                <h3 class="font-bold text-slate-800 text-lg mb-4 text-center">Trend Pencapaian Utiliti & Sisa (Kasar)</h3>
                <div class="flex-1 w-full relative min-h-[250px]">
                    <canvas id="progressChart"></canvas>
                </div>
            </div>
        </div>
    `;
    container.innerHTML = html;

    // Render Chart.js
    renderProgressChart();
}

// ==== CHART ====
function renderProgressChart() {
    const ctx = document.getElementById('progressChart');
    if (!ctx) return;
    
    if (chartInstance) {
        chartInstance.destroy();
    }
    
    const getMonthlyEmissions = (year) => {
        const d = formData.tahun[year];
        return MONTHS.map((m, i) => {
            let eElek = d.tenaga.elektrik[i] * EMISSION_FACTORS.elektrik;
            let eAir = d.air.penggunaan[i] * EMISSION_FACTORS.air;
            let eSisa = (d.sisa.buangan[i] / 1000) * EMISSION_FACTORS.sisa_buangan + 
                        (d.sisa.kitar_semula[i] / 1000) * EMISSION_FACTORS.sisa_kitar_semula;
            return eElek + eAir + eSisa;
        });
    };

    chartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: MONTHS,
            datasets: [
                {
                    label: '2025 (Garis Asas)',
                    data: getMonthlyEmissions('2025'),
                    borderColor: '#94a3b8', 
                    borderDash: [5, 5],
                    backgroundColor: 'rgba(148, 163, 184, 0.1)',
                    borderWidth: 2,
                    tension: 0.3,
                    pointRadius: 3
                },
                {
                    label: '2026 (Semasa)',
                    data: getMonthlyEmissions('2026'),
                    borderColor: '#10b981', 
                    backgroundColor: 'rgba(16, 185, 129, 0.2)',
                    borderWidth: 2,
                    fill: true,
                    tension: 0.3,
                    pointBackgroundColor: '#fff',
                    pointBorderColor: '#10b981',
                    pointRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' } },
                x: { grid: { display: false } }
            }
        }
    });
}

// ==== DATA HANDLERS ====

function updateData(section, field, value) {
    if (section === 'asas') {
        const stringFields = ['negeri', 'daerah', 'jenisFasiliti', 'namaFasiliti', 'pbt', 'jabatan'];
        formData.asas[field] = stringFields.includes(field) ? value : (Number(value) || 0);
    } else {
        formData.tahun[activeYear][section][field] = Number(value) || 0;
    }
    saveToLocalStorage();
}

function updateArrayData(section, field, index, value) {
    if (section !== 'asas') {
        formData.tahun[activeYear][section][field][index] = Number(value) || 0;
    }
    saveToLocalStorage();
}

function updateRealtimeCarbon(section) {
    const dataSection = formData.tahun[activeYear][section];
    
    if(section === 'tenaga') {
        const tKwh = sumArray(dataSection.elektrik);
        document.getElementById('tenaga-total-val').textContent = formatNumber(tKwh);
        document.getElementById('tenaga-carbon-val').textContent = formatNumber(tKwh * EMISSION_FACTORS.elektrik);
        document.getElementById('genset-carbon-val').textContent = formatNumber(dataSection.genset * EMISSION_FACTORS.genset_diesel);
    } 
    else if(section === 'air') {
        const tAir = sumArray(dataSection.penggunaan);
        document.getElementById('air-total-val').textContent = formatNumber(tAir);
        document.getElementById('air-carbon-val').textContent = formatNumber(tAir * EMISSION_FACTORS.air);
    }
    else if(section === 'sisa') {
        const tBuang = sumArray(dataSection.buangan) / 1000;
        const tKitar = sumArray(dataSection.kitar_semula) / 1000;
        document.getElementById('sisa-buangan-carbon').textContent = formatNumber(tBuang * EMISSION_FACTORS.sisa_buangan);
        document.getElementById('sisa-kitar-carbon').textContent = formatNumber(tKitar * EMISSION_FACTORS.sisa_kitar_semula);
    }
    else if(section === 'hijau') {
        const serapan = (dataSection.hutan * EMISSION_FACTORS.hutan) +
                        (dataSection.landskap * EMISSION_FACTORS.landskap) +
                        (dataSection.badan_air * EMISSION_FACTORS.badan_air) +
                        (dataSection.pokok * EMISSION_FACTORS.pokok);
        document.getElementById('hijau-serapan').innerHTML = `- ${formatNumber(serapan)} <span class="text-lg text-emerald-500">kg CO₂e</span>`;
    }
    else if(section === 'mobiliti') {
        let carbonMobiId = (dataSection.petrol * EMISSION_FACTORS.petrol) +
                           (dataSection.diesel * EMISSION_FACTORS.diesel) +
                           (dataSection.cng * EMISSION_FACTORS.cng);
        document.getElementById('mobiliti-carbon-val').innerHTML = `${formatNumber(carbonMobiId)} <span class="text-sm">kg CO₂e</span>`;
    }
}

// ==== IMPORT / EXPORT ====

function exportJSON() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(formData, null, 2));
    const filename = `LCC_Data_${(formData.asas.namaFasiliti || 'Fasiliti').replace(/\s+/g, '_')}_${formData.asas.tahunProjek}.json`;
    
    const dlAnchorElem = document.createElement('a');
    dlAnchorElem.setAttribute("href", dataStr);
    dlAnchorElem.setAttribute("download", filename);
    document.body.appendChild(dlAnchorElem);
    dlAnchorElem.click();
    dlAnchorElem.remove();
    
    showToast("Data Pendaftaran LCC berjaya dieksport.");
}

function importJSON(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = JSON.parse(e.target.result);
            
            // Check structural compatibility
            if(data.tahun && data.tahun['2025'] && data.tahun['2026']) {
                formData = { ...formData, ...data };
            } else {
                // Backward compatibility handling: if old format, dump all to 2025, clear 2026
                formData.asas = { ...formData.asas, ...data.asas };
                formData.tahun['2025'] = { 
                    tenaga: data.tenaga || formData.tahun['2025'].tenaga, 
                    air: data.air || formData.tahun['2025'].air, 
                    sisa: data.sisa || formData.tahun['2025'].sisa, 
                    hijau: data.hijau || formData.tahun['2025'].hijau, 
                    mobiliti: data.mobiliti || formData.tahun['2025'].mobiliti
                };
            }
            
            if(!formData.asas.negeri) formData.asas.negeri = '';
            if(!formData.asas.daerah) formData.asas.daerah = '';
            
            renderSection(activeSection);
            showToast("Data Fasiliti berjaya dimuat naik!");
        } catch (error) {
            alert('Ralat membaca fail JSON. Pastikan format betul.');
        }
    };
    reader.readAsText(file);
    event.target.value = ''; // Reset import
}

function showToast(msg) {
    const t = document.getElementById('toast');
    document.getElementById('toast-msg').textContent = msg;
    t.classList.remove('opacity-0', 'translate-y-20');
    
    setTimeout(() => {
        t.classList.add('opacity-0', 'translate-y-20');
    }, 3000);
}

// ==== MUAT TURUN MS WORD ====

function exportWord() {
    // Generate simple MS Word compatible XML HTML
    const htmlHeader = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
        <meta charset='utf-8'>
        <title>Laporan LCC 2030</title>
        <style>
            body { font-family: 'Calibri', sans-serif; }
            h1 { color: #047857; text-align: center; }
            h2 { color: #1e293b; border-bottom: 1px solid #047857; padding-bottom: 4px; margin-top: 24px; }
            table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
            th, td { border: 1px solid #cbd5e1; padding: 8px; text-align: left; }
            th { background-color: #f1f5f9; }
            .right { text-align: right; }
            .bg-gray { background-color: #f8fafc; }
        </style>
    </head>
    <body>`;

    const htmlFooter = `</body></html>`;

    // Kiraan Total
    let d2025 = calculateEmissions('2025');
    let d2026 = calculateEmissions('2026');
    let varNet = calculateVariance(d2025.net, d2026.net);

    let content = `
        <h1>Laporan Karbon Fasiliti LCC 2030</h1>
        
        <h2>1. Profil Fasiliti</h2>
        <table>
            <tr><th>Negeri</th><td>${formData.asas.negeri || '-'}</td></tr>
            <tr><th>Daerah</th><td>${formData.asas.daerah || '-'}</td></tr>
            <tr><th>PBT Kawasan</th><td>${formData.asas.pbt || '-'}</td></tr>
            <tr><th>Jenis Fasiliti</th><td>${formData.asas.jenisFasiliti || '-'}</td></tr>
            <tr><th>Nama Fasiliti</th><td>${formData.asas.namaFasiliti || '-'}</td></tr>
            <tr><th>Tahun Asas</th><td>${formData.asas.tahunAsas || '-'}</td></tr>
            <tr><th>Tahun Semasa</th><td>${formData.asas.tahunProjek || '-'}</td></tr>
            <tr><th>Populasi Kakitangan</th><td>${formData.asas.populasi || '-'}</td></tr>
        </table>

        <h2>2. Perbandingan Inventori Pelepasan (kg CO₂e)</h2>
        <table>
            <tr><th>Kategori / Sektor</th><th class="right">Asas (2025)</th><th class="right">Semasa (2026)</th></tr>
            <tr><td>Tenaga (Elektrik & Genset)</td><td class="right">${formatNumber(d2025.t_energy)}</td><td class="right">${formatNumber(d2026.t_energy)}</td></tr>
            <tr><td>Kewangan Air Bersih</td><td class="right">${formatNumber(d2025.t_air)}</td><td class="right">${formatNumber(d2026.t_air)}</td></tr>
            <tr><td>Sisa Pepejal (Pelupusan & Kitar Semula)</td><td class="right">${formatNumber(d2025.t_waste)}</td><td class="right">${formatNumber(d2026.t_waste)}</td></tr>
            <tr><td>Mobiliti & Pengangkutan Skop 1</td><td class="right">${formatNumber(d2025.t_mobility)}</td><td class="right">${formatNumber(d2026.t_mobility)}</td></tr>
            <tr class="bg-gray"><th>JUMLAH PELEPASAN KASAR (GROSS)</th><th class="right">${formatNumber(d2025.gross)}</th><th class="right">${formatNumber(d2026.gross)}</th></tr>
        </table>
        
        <h2>3. Serapan Karbon Positif (Sink)</h2>
        <table>
            <tr><th>Lanskap & Ekologi</th><th class="right">Asas 2025</th><th class="right">Semasa 2026</th></tr>
            <tr><td>Kawasan Hijau, Air & Pokok</td><td class="right">- ${formatNumber(d2025.sink)}</td><td class="right">- ${formatNumber(d2026.sink)}</td></tr>
        </table>

        <h2>4. Pelepasan Karbon Bersih (NET) & Keputusan Pencapaian</h2>
        <p style="font-size: 16px;">Jejak Karbon Bersih 2025: <b>${formatNumber(d2025.net)} kg CO₂e</b></p>
        <p style="font-size: 16px;">Jejak Karbon Bersih 2026: <b>${formatNumber(d2026.net)} kg CO₂e</b></p>
        <p style="font-size: 18px; color: ${varNet.isReduction ? '#047857' : '#b41c1c'};">Perbezaan Variance: <b>${varNet.isReduction ? '-' : '+'}${formatNumber(varNet.pct, 1)}% (${varNet.isReduction ? 'Pengurangan Berjaya' : 'Peningkatan Gagal'})</b></p>
        
        <p style="color:#64748b; font-size:12px; margin-top:40px;"><i>Dijana secara automatik oleh Borang Interaktif LCC 2030 SPA.</i></p>
    `;

    const fullBlobHTML = htmlHeader + content + htmlFooter;
    const blob = new Blob(['\ufeff', fullBlobHTML], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    
    const filename = `Laporan_LCC_${(formData.asas.namaFasiliti || 'Fasiliti').replace(/\s+/g, '_')}_${formData.asas.tahunProjek}.doc`;
    
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    showToast("Dokumen Microsoft Word berjaya dimuat turun.");
}

// ==== HANTAR KE GOOGLE SHEETS ====

async function sendToGoogleSheet() {
    if (GOOGLE_SHEET_WEB_APP_URL === "") {
        alert("PERHATIAN: Fungsi belum dipautkan!\n\nSila bina satu skrip Google Apps (rujuk panduan) dan letakkan URL tersebut di dalam fail app.js pada pembolehubah GOOGLE_SHEET_WEB_APP_URL.");
        return;
    }

    const btn = document.getElementById('btn-google');
    const originalHtml = btn.innerHTML;
    // Tukar butang kepada gaya loading
    btn.innerHTML = `<i class="ph ph-spinner animate-spin text-lg relative z-10"></i> <span class="relative z-10">Mengirim...</span>`;
    btn.classList.add('opacity-80', 'cursor-not-allowed');
    btn.disabled = true;

    try {
        let d2025 = calculateEmissions('2025');
        let d2026 = calculateEmissions('2026');
        let varNet = calculateVariance(d2025.net, d2026.net);
        
        // Bina payload yang leper (flat) untuk Google Sheet supaya mudah dipapar pada lajur
        const sheetPayload = {
            timestamp: new Date().toISOString(),
            negeri: formData.asas.negeri,
            daerah: formData.asas.daerah,
            jenisFasiliti: formData.asas.jenisFasiliti,
            namaFasiliti: formData.asas.namaFasiliti,
            pbt: formData.asas.pbt,
            populasi: formData.asas.populasi,
            
            // 2025 Baseline
            asas_tenaga: d2025.t_energy,
            asas_air: d2025.t_air,
            asas_sisa: d2025.t_waste,
            asas_mobiliti: d2025.t_mobility,
            asas_kasar: d2025.gross,
            asas_sink: d2025.sink,
            asas_net: d2025.net,
            
            // 2026 Pencapaian
            semasa_tenaga: d2026.t_energy,
            semasa_air: d2026.t_air,
            semasa_sisa: d2026.t_waste,
            semasa_mobiliti: d2026.t_mobility,
            semasa_kasar: d2026.gross,
            semasa_sink: d2026.sink,
            semasa_net: d2026.net,

            // Pencapaian Reduce
            peratus_pengurangan: (varNet.isReduction ? -1 : 1) * varNet.pct,
            pengurangan_kg: (varNet.isReduction ? -1 : 1) * varNet.diff,
            
            // Data mentah lengkap berbentuk string untuk back-up
            dataMentahJSON: JSON.stringify(formData)
        };

        const response = await fetch(GOOGLE_SHEET_WEB_APP_URL, {
            method: 'POST',
            mode: 'no-cors', // Atasi isu CORS dari browser client ke Google
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(sheetPayload)
        });
        
        showToast("Laporan anda telah selamat direkodkan ke pelayan kementerian!");
        
    } catch (error) {
        console.error(error);
        alert("Gagal menghantar laporan. Sila periksa capaian internet anda.");
    } finally {
        btn.innerHTML = originalHtml;
        btn.classList.remove('opacity-80', 'cursor-not-allowed');
        btn.disabled = false;
    }
}
