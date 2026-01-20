// 심의자료 자동화 웹 UI 스크립트 v2.1 - SSE 실시간 진행상황

document.addEventListener('DOMContentLoaded', function () {
    // DOM 요소
    const loadAllBtn = document.getElementById('loadAllBtn');
    const filterSection = document.getElementById('filterSection');
    const filterInput = document.getElementById('filterInput');
    const selectAllBtn = document.getElementById('selectAllBtn');
    const deselectAllBtn = document.getElementById('deselectAllBtn');
    const materialsSection = document.getElementById('materialsSection');
    const materialsGrid = document.getElementById('materialsGrid');
    const selectedCount = document.getElementById('selectedCount');
    const totalCount = document.getElementById('totalCount');
    const actionSection = document.getElementById('actionSection');
    const generateBtn = document.getElementById('generateBtn');
    const progressSection = document.getElementById('progressSection');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const statusSection = document.getElementById('statusSection');
    const statusMessage = document.getElementById('statusMessage');

    let allMaterials = {};
    let selectedMaterials = new Set();

    // 전체 소재 불러오기
    loadAllBtn.addEventListener('click', loadAllMaterials);

    // 필터 입력
    filterInput.addEventListener('input', filterMaterials);

    // 전체 선택/해제
    selectAllBtn.addEventListener('click', () => toggleAll(true));
    deselectAllBtn.addEventListener('click', () => toggleAll(false));

    // PPT 생성
    generateBtn.addEventListener('click', generatePPT);

    async function loadAllMaterials() {
        const btnText = loadAllBtn.querySelector('.btn-text');
        const btnLoading = loadAllBtn.querySelector('.btn-loading');
        btnText.style.display = 'none';
        btnLoading.style.display = 'inline-flex';
        loadAllBtn.disabled = true;

        showStatus('소재 목록을 불러오는 중...', 'loading');

        try {
            const response = await fetch('/api/all-materials');
            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.detail || '소재 목록을 불러올 수 없습니다.');
            }

            allMaterials = data.details;
            displayMaterials(data.materials, data.details);

            filterSection.style.display = 'block';
            materialsSection.style.display = 'block';
            actionSection.style.display = 'block';

            showStatus(`${data.count}개 소재를 불러왔습니다. 생성할 소재를 선택하세요.`, 'success');

        } catch (error) {
            showStatus(error.message, 'error');
        } finally {
            btnText.style.display = 'inline';
            btnLoading.style.display = 'none';
            loadAllBtn.disabled = false;
        }
    }

    function displayMaterials(materials, details) {
        materialsGrid.innerHTML = '';
        totalCount.textContent = materials.length;

        materials.forEach(name => {
            const sizes = details[name] || [];
            const item = document.createElement('div');
            item.className = 'material-item';
            item.dataset.name = name;

            item.innerHTML = `
                <input type="checkbox" id="mat_${name}" value="${name}">
                <span class="material-name">${name}</span>
                <span class="material-count">${sizes.length}개</span>
            `;

            const checkbox = item.querySelector('input');
            checkbox.addEventListener('change', () => {
                if (checkbox.checked) {
                    selectedMaterials.add(name);
                    item.classList.add('selected');
                } else {
                    selectedMaterials.delete(name);
                    item.classList.remove('selected');
                }
                updateSelectedCount();
            });

            item.addEventListener('click', (e) => {
                if (e.target.tagName !== 'INPUT') {
                    checkbox.checked = !checkbox.checked;
                    checkbox.dispatchEvent(new Event('change'));
                }
            });

            materialsGrid.appendChild(item);
        });

        updateSelectedCount();
    }

    function filterMaterials() {
        const filter = filterInput.value.toLowerCase();
        const items = materialsGrid.querySelectorAll('.material-item');

        items.forEach(item => {
            const name = item.dataset.name.toLowerCase();
            if (name.includes(filter)) {
                item.classList.remove('hidden');
            } else {
                item.classList.add('hidden');
            }
        });
    }

    function toggleAll(select) {
        const items = materialsGrid.querySelectorAll('.material-item:not(.hidden)');
        items.forEach(item => {
            const checkbox = item.querySelector('input');
            checkbox.checked = select;

            if (select) {
                selectedMaterials.add(item.dataset.name);
                item.classList.add('selected');
            } else {
                selectedMaterials.delete(item.dataset.name);
                item.classList.remove('selected');
            }
        });
        updateSelectedCount();
    }

    function updateSelectedCount() {
        selectedCount.textContent = selectedMaterials.size;
        generateBtn.disabled = selectedMaterials.size === 0;
    }

    async function generatePPT() {
        if (selectedMaterials.size === 0) {
            showStatus('소재를 선택해주세요.', 'error');
            return;
        }

        // 버튼 상태 변경
        const btnText = generateBtn.querySelector('.btn-text');
        const btnLoading = generateBtn.querySelector('.btn-loading');
        btnText.style.display = 'none';
        btnLoading.style.display = 'inline-flex';
        generateBtn.disabled = true;

        // 진행상황 표시
        progressSection.style.display = 'block';
        progressFill.style.width = '0%';
        progressText.textContent = '준비 중...';

        hideStatus();

        const materialsParam = Array.from(selectedMaterials).join(',');

        try {
            // SSE 연결
            const eventSource = new EventSource(`/api/generate-sse?materials=${encodeURIComponent(materialsParam)}`);

            eventSource.onmessage = function (event) {
                const data = JSON.parse(event.data);

                if (data.type === 'progress') {
                    // 진행상황 업데이트
                    progressFill.style.width = `${data.percent}%`;
                    progressText.textContent = `[${data.step}] ${data.detail || ''} (${data.percent}%)`;
                } else if (data.type === 'complete') {
                    // 완료 - 다운로드
                    eventSource.close();
                    progressFill.style.width = '100%';
                    progressText.textContent = '✅ 완료! 다운로드 중...';

                    // 다운로드 링크 생성
                    window.location.href = `/api/download/${data.token}`;

                    showStatus('✅ PPT 파일이 다운로드되었습니다!', 'success');
                    resetButton();
                } else if (data.type === 'error') {
                    eventSource.close();
                    showStatus(`❌ 오류: ${data.message}`, 'error');
                    progressText.textContent = '오류 발생';
                    resetButton();
                }
            };

            eventSource.onerror = function (e) {
                eventSource.close();
                showStatus('서버 연결 오류가 발생했습니다.', 'error');
                progressText.textContent = '연결 오류';
                resetButton();
            };

        } catch (error) {
            showStatus(error.message, 'error');
            progressText.textContent = '오류 발생';
            resetButton();
        }

        function resetButton() {
            btnText.style.display = 'inline';
            btnLoading.style.display = 'none';
            generateBtn.disabled = false;
        }
    }

    function showStatus(message, type) {
        statusMessage.textContent = message;
        statusMessage.className = `status-message ${type}`;
        statusSection.style.display = 'block';
    }

    function hideStatus() {
        statusSection.style.display = 'none';
    }
});
