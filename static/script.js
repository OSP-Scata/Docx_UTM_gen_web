document.addEventListener('DOMContentLoaded', () => {

    // ==========================================
    // 1. ГЛОБАЛЬНОЕ ОБЪЯВЛЕНИЕ ВСЕХ ПЕРЕМЕННЫХ
    // ==========================================
    // Элементы модального окна и текстов
    const helpModal = document.getElementById('helpModal');
    const openHelpBtn = document.getElementById('openHelpBtn');
    const closeHelpBtn = document.getElementById('closeHelpBtn');
    const helpSingleText = document.getElementById('helpSingleText');
    const helpBatchText = document.getElementById('helpBatchText');
    
    // Элементы вкладок и форм
    const tabSingle = document.getElementById('tabSingle');
    const tabBatch = document.getElementById('tabBatch');
    const singleBlock = document.getElementById('singleFormBlock');
    const batchBlock = document.getElementById('batchFormBlock');
    const statusDiv = document.getElementById('status');

    // ==========================================
    // 2. КОНТЕКСТНАЯ ЛОГИКА ИНСТРУКЦИИ
    // ==========================================
    if (openHelpBtn && helpModal && closeHelpBtn && helpSingleText && helpBatchText && singleBlock) {
        openHelpBtn.onclick = function(e) {
            e.preventDefault();
            
            // Проверяем, какой блок сейчас отображается на экране
            if (singleBlock.classList.contains('active-content')) {
                helpSingleText.style.display = 'block';
                helpBatchText.style.display = 'none';
            } else {
                helpSingleText.style.display = 'none';
                helpBatchText.style.display = 'block';
            }
            
            helpModal.style.display = 'flex';
        };

        closeHelpBtn.onclick = function(e) {
            e.preventDefault();
            helpModal.style.display = 'none';
        };

        helpModal.onclick = function(e) {
            if (e.target === helpModal) {
                helpModal.style.display = 'none';
            }
        };
    }

    // ==========================================
    // 3. ЛОГИКА ПЕРЕКЛЮЧЕНИЯ ВКЛАДОК (ТАБОВ)
    // ==========================================
    function switchTab(activeTab, activeBlock, inactiveTab, inactiveBlock) {
        if (statusDiv) statusDiv.style.display = 'none';
        
        activeTab.classList.add('active');
        activeBlock.classList.add('active-content');
        
        inactiveTab.classList.remove('active');
        inactiveBlock.classList.remove('active-content');
    }

    if (tabSingle && tabBatch && singleBlock && batchBlock) {
        tabSingle.onclick = function(e) {
            e.preventDefault();
            switchTab(tabSingle, singleBlock, tabBatch, batchBlock);
        };

        tabBatch.onclick = function(e) {
            e.preventDefault();
            switchTab(tabBatch, batchBlock, tabSingle, singleBlock);
        };
    }

    // ==========================================
    // 4. ОТОБРАЖЕНИЕ ИМЕН ФАЙЛОВ ПРИ ЗАГРУЗКЕ
    // ==========================================
    const setupFileInput = (inputId, fileNameId, prefixText) => {
        const input = document.getElementById(inputId);
        const nameDiv = document.getElementById(fileNameId);
        if (input && nameDiv) {
            input.onchange = function(e) {
                if (e.target.files.length > 0) {
                    nameDiv.innerText = prefixText + e.target.files[0].name; // ИСПРАВЛЕНО: .files[0].name для корректного чтения имени файла
                    nameDiv.style.display = 'block';
                }
            };
        }
    };

    setupFileInput('fileInput', 'fileName', 'Выбран файл: ');
    setupFileInput('excelInput', 'excelFileName', 'Выбран манифест: ');
    setupFileInput('zipInput', 'zipFileName', 'Выбран архив: ');

    // ==========================================
    // 5. ОТПРАВКА ФОРМЫ 1 (ОДИНОЧНАЯ)
    // ==========================================
    const utmForm = document.getElementById('utmForm');
    if (utmForm) {
        utmForm.onsubmit = async function(e) {
            e.preventDefault();
            const submitBtn = document.getElementById('submitBtn');
            if (!statusDiv || !submitBtn) return;

            statusDiv.style.display = 'none';
            submitBtn.innerText = "Обработка и упаковка...";
            submitBtn.disabled = true;

            try {
                const response = await fetch('/generate-utm/', {
                    method: 'POST',
                    body: new FormData(this)
                });
                if (!response.ok) {
                    const errData = await response.json();
                    throw new Error(errData.detail || 'Ошибка сервера');
                }
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = "utm_packed.zip";
                document.body.appendChild(a);
                a.click();
                a.remove();

                statusDiv.className = "success";
                statusDiv.innerText = "Готово! Архив успешно скачан.";
                statusDiv.style.display = 'block';
            } catch (error) {
                statusDiv.className = "error";
                statusDiv.innerText = "Ошибка: " + error.message;
                statusDiv.style.display = 'block';
            } finally {
                submitBtn.innerText = "Сгенерировать и скачать пакет";
                submitBtn.disabled = false;
            }
        };
    }

    // ==========================================
    // 6. ОТПРАВКА ФОРМЫ 2 (МАССОВАЯ)
    // ==========================================
    const batchUtmForm = document.getElementById('batchUtmForm');
    if (batchUtmForm) {
        batchUtmForm.onsubmit = async function(e) {
            e.preventDefault();
            const batchSubmitBtn = document.getElementById('batchSubmitBtn');
            if (!statusDiv || !batchSubmitBtn) return;

            statusDiv.style.display = 'none';
            batchSubmitBtn.innerText = "Распаковка и разметка...";
            batchSubmitBtn.disabled = true;

            try {
                const response = await fetch('/generate-utm-batch/', {
                    method: 'POST',
                    body: new FormData(this)
                });
                if (!response.ok) {
                    const errData = await response.json();
                    throw new Error(errData.detail || 'Ошибка пакетной обработки');
                }
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = "mass_utm_packed.zip";
                document.body.appendChild(a);
                a.click();
                a.remove();

                statusDiv.className = "success";
                statusDiv.innerText = "Пакетная обработка завершена! Архив скачан.";
                statusDiv.style.display = 'block';
            } catch (error) {
                statusDiv.className = "error";
                statusDiv.innerText = "Ошибка: " + error.message;
                statusDiv.style.display = 'block';
            } finally {
                batchSubmitBtn.innerText = "Обработать и скачать пакет";
                batchSubmitBtn.disabled = false;
            }
        };
    }
     // выпадающий список utm_medium
    const mediumSelect = document.getElementById('medium');
    const customMediumInput = document.getElementById('custom_medium');
    
    if (mediumSelect && customMediumInput) {
        mediumSelect.addEventListener('change', (e) => {
            if (e.target.value === 'custom') {
                customMediumInput.style.display = 'block';
                customMediumInput.required = true;
            } else {
                customMediumInput.style.display = 'none';
                customMediumInput.required = false;
            }
        });
    }

    // спойлер доп.параметров
    const spoilerBtn = document.getElementById('spoilerBtn');
    const spoilerContent = document.getElementById('spoilerContent');
    
    if (spoilerBtn && spoilerContent) {
        spoilerBtn.onclick = function(e) {
            e.preventDefault();
            e.stopPropagation();
            if (spoilerContent.style.display === 'none' || spoilerContent.style.display === '') {
                spoilerContent.style.display = 'block';
                spoilerBtn.innerText = "⚙️ Дополнительные параметры (необязательно) ▴";
            } else {
                spoilerContent.style.display = 'none';
                spoilerBtn.innerText = "⚙️ Дополнительные параметры (необязательно) ▾";
            }
        };
    }
});