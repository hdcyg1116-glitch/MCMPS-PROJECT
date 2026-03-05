document.addEventListener('DOMContentLoaded', function () {
    console.log("공정 관리 시스템 UI 로드 완료");
    fetchData();
    updateClock();
    setInterval(updateClock, 1000); // 1초마다 시계 업데이트

    // 차트 인스턴스 저장용 변수
    window.overallChartInstance = null;

    // 드래그 앤 드롭 핸들러
    const dropArea = document.getElementById('drop-area');
    const fileElem = document.getElementById('fileElem');

    // 클릭 시 파일 선택창 띄우기 (Label for로 대체되어 필요 없음)

    // 파일 선택창에서 파일 선택 시 처리
    fileElem.addEventListener('change', function (e) {
        handleFiles(this.files);
    });

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    dropArea.addEventListener('drop', handleDrop, false);

    function handleDrop(e) {
        let dt = e.dataTransfer;
        let files = dt.files;
        handleFiles(files);
    }

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];

            // 파일 확장자 검사 (창 로딩 속도 최적화를 위해 HTML accept 속성 대신 JS에서 처리)
            if (!file.name.match(/\.(xlsx|xls)$/i)) {
                alert("엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.");
                return;
            }

            console.log("파일 업로드 시작:", file.name);

            // 로딩 표시
            document.getElementById('last-updated').innerText = "업데이트 중...";

            const formData = new FormData();
            formData.append('file', file);

            fetch('/api/upload', {
                method: 'POST',
                body: formData
            })
                .then(response => response.json())
                .then(data => {
                    if (data.error) {
                        alert("업로드 실패: " + data.error);
                        document.getElementById('last-updated').innerText = "업로드 실패";
                    } else if (data.length === 0) {
                        alert("분석된 데이터가 없습니다. 엑셀 파일 형식을 확인해주세요.");
                        document.getElementById('last-updated').innerText = "데이터 없음";
                    } else {
                        console.log("파일 분석 완료, 화면 업데이트");

                        // 필터 및 검색어 초기화 (새 데이터가 가려지지 않게 함)
                        const searchInput = document.getElementById('global-search');
                        if (searchInput) searchInput.value = '';
                        window.columnFilters = {};

                        window.allProductionData = data;
                        populateSectionFilter(data);
                        filterData(); // 자동 정렬 적용
                        renderDashboardSummaries(data);
                        document.getElementById('last-updated').innerText = `최근 파일: ${file.name} (${new Date().toLocaleTimeString()})`;
                    }
                })
                .catch(error => {
                    console.error('Error uploading file:', error);
                    alert("파일 전송 중 오류가 발생했습니다.");
                });
        }
    }
});

// 전역 데이터 저장용
window.allProductionData = [];
window.columnFilters = {}; // 컬럼별 필터 전역 상태
window.currentSort = { key: null, direction: 'asc' }; // 정렬 상태 기록

function fetchData() {
    fetch('/api/data')
        .then(response => response.json())
        .then(data => {
            window.allProductionData = data;
            populateSectionFilter(data);
            populateColumnFilters(data);
            filterData(); // 필터 적용 및 테이블 렌더링
            renderDashboardSummaries(data);
            document.getElementById('last-updated').innerText = "최근 업데이트: " + new Date().toLocaleTimeString();
        })
        .catch(error => console.error('Error fetching data:', error));
}

function populateSectionFilter(data) {
    const filterSelect = document.getElementById('filter-section');
    if (!filterSelect) return;

    const currentVal = filterSelect.value;
    const sections = [...new Set(data.map(item => item.section).filter(s => s && s !== '-'))].sort();

    // 초기화 (전체 옵션은 유지)
    filterSelect.innerHTML = '<option value="all">전체 (직)</option>';
    sections.forEach(sec => {
        const opt = document.createElement('option');
        opt.value = sec;
        opt.innerText = sec;
        filterSelect.appendChild(opt);
    });

    // 이전 선택값 유지
    if (sections.includes(currentVal)) {
        filterSelect.value = currentVal;
    }
}

// 필터 상태 변경 
// 이제 columnFilters는 { 'key': ['A', 'B'] } 형태의 배열을 가집니다. (빈 배열이면 필터 없음)

let activeFilterPopup = null;

function populateColumnFilters(data) {
    const headers = document.querySelectorAll('#production-table th');

    // 바탕화면 클릭 시 팝업 닫기 이벤트 등록
    document.addEventListener('mousedown', function (e) {
        if (activeFilterPopup && !activeFilterPopup.contains(e.target)) {
            // 클릭한 곳이 헤더 필터 아이콘인지 확인 (아이콘 클릭 시 닫히고 다시 열리는 것 방지)
            const isFilterIcon = e.target.closest('.header-filter-icon');
            if (isFilterIcon) {
                const popupKey = activeFilterPopup.getAttribute('data-key');
                const clickedKey = isFilterIcon.closest('th').getAttribute('data-key');
                if (popupKey === clickedKey) return;
            }
            closeFilterPopup();
        }
    });

    headers.forEach(th => {
        const key = th.getAttribute('data-key');
        if (!key) return;

        // 기존 select 엘리먼트가 남아있다면 제거
        const oldSelect = th.querySelector('select');
        if (oldSelect) oldSelect.remove();

        // 아이콘 그룹 컨테이너 확인/생성
        let actionsContainer = th.querySelector('.header-actions');
        if (!actionsContainer) {
            actionsContainer = document.createElement('div');
            actionsContainer.className = 'header-actions';
            th.appendChild(actionsContainer);
        }

        // 필터 아이콘 버튼 생성
        let filterIcon = actionsContainer.querySelector('.header-filter-icon');
        if (!filterIcon) {
            filterIcon = document.createElement('span');
            filterIcon.className = 'header-filter-icon';
            filterIcon.innerHTML = '▼';
            filterIcon.title = "필터 적용";

            filterIcon.addEventListener('click', function (e) {
                e.stopPropagation(); // 헤더 정렬 방지
                toggleFilterPopup(th, key);
            });
            actionsContainer.appendChild(filterIcon);
        }

        if (!(key in window.columnFilters)) {
            window.columnFilters[key] = []; // 배열로 초기화
        }
    });
}

function toggleFilterPopup(th, key) {
    if (activeFilterPopup) {
        const currentKey = activeFilterPopup.getAttribute('data-key');
        closeFilterPopup();
        if (currentKey === key) return; // 같은 아이콘을 클릭한 경우 닫기만 하고 종료
    }

    createFilterPopup(th, key);
}

function closeFilterPopup() {
    if (activeFilterPopup) {
        activeFilterPopup.remove();
        activeFilterPopup = null;
    }
}

function createFilterPopup(th, key) {
    const popup = document.createElement('div');
    popup.className = 'excel-filter-menu';
    popup.setAttribute('data-key', key);

    // 현재 필터 상태(복사본)
    let tempFilterSelection = [...(window.columnFilters[key] || [])];

    // 1. 다른 컬럼들의 필터 조건이 적용된 데이터 추출 (종속 필터 구현)
    let dataForThisColumn = window.allProductionData || [];
    for (const k in window.columnFilters) {
        if (k !== key && window.columnFilters[k] && window.columnFilters[k].length > 0) {
            dataForThisColumn = dataForThisColumn.filter(item => window.columnFilters[k].includes(String(item[k])));
        }
    }

    // 존재하는 고유값 추출 (전부 문자열로 변환하여 타입 불일치 방지)
    const uniqueValues = [...new Set(dataForThisColumn.map(item => {
        let itemVal = item[key];
        if (itemVal === undefined || itemVal === null) itemVal = '';
        return String(itemVal);
    }))].sort();

    // 이전에 선택된 항목들이 더 이상 유효하지 않은 경우, 상태 정리
    tempFilterSelection = tempFilterSelection.map(String).filter(val => uniqueValues.includes(val));

    // HTML 구성
    const hasFilter = window.columnFilters[key] && window.columnFilters[key].length > 0;

    let html = `
        <div class="excel-filter-options">
            <div class="excel-filter-item sort-asc" data-sort="asc">
                <span style="font-size:1.1rem">A↓</span> 텍스트 오름차순 정렬
            </div>
            <div class="excel-filter-item sort-desc" data-sort="desc">
                <span style="font-size:1.1rem">Z↓</span> 텍스트 내림차순 정렬
            </div>
            <div class="excel-filter-item clear-filter ${hasFilter ? '' : 'disabled'}" style="margin-top:0.3rem">
                <span style="font-size:1.1rem; color:red">✗</span> "${th.innerText.replace('▼', '').trim()}" 에서 필터 해제
            </div>
        </div>
        <div class="excel-filter-search-container">
            <input type="text" class="excel-filter-search" placeholder="검색">
        </div>
        <div class="excel-filter-list" id="filter-list-${key}">
            <label class="excel-checkbox-item">
                <input type="checkbox" id="selectAll-${key}"> (모두 선택)
            </label>
    `;

    uniqueValues.forEach(val => {
        const isChecked = tempFilterSelection.length === 0 || tempFilterSelection.includes(val);
        html += `
            <label class="excel-checkbox-item data-item">
                <input type="checkbox" value="${val}" ${isChecked ? 'checked' : ''}> ${val}
            </label>
        `;
    });

    html += `
        </div>
        <div class="excel-filter-actions">
            <button class="excel-filter-btn" id="cancelFilterBtn">취소</button>
            <button class="excel-filter-btn primary" id="applyFilterBtn">확인</button>
        </div>
    `;

    popup.innerHTML = html;
    document.body.appendChild(popup);
    activeFilterPopup = popup;

    // 위치 설정
    const rect = th.getBoundingClientRect();
    popup.style.top = (rect.bottom + window.scrollY + 2) + 'px';
    popup.style.left = (rect.left + window.scrollX) + 'px';

    // (이벤트 연결)
    attachPopupEvents(popup, key, th, tempFilterSelection, uniqueValues);
}

function attachPopupEvents(popup, key, th, tempFilterSelection, uniqueValues) {
    const checkboxes = popup.querySelectorAll('.data-item input[type="checkbox"]');
    const selectAllCheckbox = popup.querySelector(`#selectAll-${key}`);
    const searchInput = popup.querySelector('.excel-filter-search');

    // 초기 '모두 선택' 체크 확인
    let checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;
    selectAllCheckbox.checked = (checkedCount === checkboxes.length) || tempFilterSelection.length === 0;

    // 전체 선택/해제 로직
    selectAllCheckbox.addEventListener('change', function () {
        const visibleCheckboxes = popup.querySelectorAll('.data-item:not([style*="display: none"]) input[type="checkbox"]');
        visibleCheckboxes.forEach(cb => {
            cb.checked = this.checked;
        });
    });

    // 개별 체크박스 로직
    checkboxes.forEach(cb => {
        cb.addEventListener('change', function () {
            const visibleCheckboxes = popup.querySelectorAll('.data-item:not([style*="display: none"]) input[type="checkbox"]');
            const allChecked = Array.from(visibleCheckboxes).every(c => c.checked);
            selectAllCheckbox.checked = allChecked;
        });
    });

    // 검색어 필터링
    searchInput.addEventListener('input', function () {
        const text = this.value.toLowerCase();
        const labels = popup.querySelectorAll('.data-item');
        labels.forEach(label => {
            const val = label.innerText.toLowerCase();
            if (val.includes(text)) {
                label.style.display = 'flex';
            } else {
                label.style.display = 'none';
            }
        });

        // 보이는 항목만으로 selectAll 상태 갱신
        const visibleCheckboxes = popup.querySelectorAll('.data-item:not([style*="display: none"]) input[type="checkbox"]');
        if (visibleCheckboxes.length > 0) {
            selectAllCheckbox.checked = Array.from(visibleCheckboxes).every(c => c.checked);
        } else {
            selectAllCheckbox.checked = false;
        }
    });

    // 엔터키 단축키 지원
    searchInput.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
            const applyBtn = popup.querySelector('#applyFilterBtn');
            if (applyBtn) applyBtn.click();
        }
    });

    // 정렬 클릭
    const sortAsc = popup.querySelector('.sort-asc');
    const sortDesc = popup.querySelector('.sort-desc');

    sortAsc.addEventListener('click', () => {
        window.currentSort.key = key;
        window.currentSort.direction = 'asc';
        closeFilterPopup();
        filterData();
    });

    sortDesc.addEventListener('click', () => {
        window.currentSort.key = key;
        window.currentSort.direction = 'desc';
        closeFilterPopup();
        filterData();
    });

    // 필터 해제 클릭
    const clearFilter = popup.querySelector('.clear-filter');
    if (!clearFilter.classList.contains('disabled')) {
        clearFilter.addEventListener('click', () => {
            window.columnFilters[key] = [];
            closeFilterPopup();
            filterData();
        });
    }

    // 확인 버튼
    popup.querySelector('#applyFilterBtn').addEventListener('click', () => {
        // 보이는 체크박스들의 상태만 가져옵니다. (또는 전체를 기준으로 해도 무방하나 사용자 경험상 전체 기준 처리)
        const updatedSelection = [];
        let isAllChecked = true;

        checkboxes.forEach(cb => {
            if (cb.checked) {
                updatedSelection.push(cb.value);
            } else {
                isAllChecked = false;
            }
        });

        // 전부 선택되어 있으면 필터 해제로 간주
        if (isAllChecked || updatedSelection.length === uniqueValues.length || updatedSelection.length === 0) {
            window.columnFilters[key] = [];
        } else {
            window.columnFilters[key] = updatedSelection;
        }

        closeFilterPopup();
        filterData();
    });

    // 취소 버튼
    popup.querySelector('#cancelFilterBtn').addEventListener('click', () => {
        closeFilterPopup();
    });

    // 내부 클릭 시 팝업 닫히지 않도록
    popup.addEventListener('mousedown', e => e.stopPropagation());
    popup.addEventListener('click', e => e.stopPropagation());
}

function handleColumnFilterChange(e) {
    const select = e.target;
    const th = select.closest('th');
    const key = th.getAttribute('data-key');

    window.columnFilters[key] = select.value;
    filterData();
}

function filterData() {
    const searchInput = document.getElementById('global-search');
    const searchQuery = searchInput ? searchInput.value.toLowerCase().trim() : '';

    let filtered = window.allProductionData;

    // 1. 글로벌 검색 필터 적용 (모든 컬럼 대상)
    if (searchQuery) {
        filtered = filtered.filter(item => {
            // 모든 속성 값을 하나의 문자열로 합쳐서 검색 (undefined/null 제외)
            return Object.values(item).some(val =>
                val !== null && val !== undefined && String(val).toLowerCase().includes(searchQuery)
            );
        });
    }

    // 2. 컬럼별 배열 필터 적용 (다중 선택 대응)
    for (const key in window.columnFilters) {
        const filters = window.columnFilters[key];
        if (filters && filters.length > 0) {
            filtered = filtered.filter(item => {
                let itemVal = item[key];
                if (itemVal === undefined || itemVal === null) itemVal = '';
                return filters.includes(String(itemVal));
            });
        }
    }

    // 3. 정렬 적용
    if (window.currentSort.key) {
        filtered.sort((a, b) => {
            let valA = a[window.currentSort.key];
            let valB = b[window.currentSort.key];

            if (valA === undefined || valA === null) valA = '';
            if (valB === undefined || valB === null) valB = '';

            valA = String(valA);
            valB = String(valB);

            // 두 값이 모두 완전히 숫자 형태일 때만 숫자 정렬로 처리
            const isNumericA = !isNaN(Number(valA)) && valA.trim() !== '' && valA.trim() !== '-';
            const isNumericB = !isNaN(Number(valB)) && valB.trim() !== '' && valB.trim() !== '-';

            if (isNumericA && isNumericB) {
                const numA = Number(valA);
                const numB = Number(valB);
                return window.currentSort.direction === 'asc' ? numA - numB : numB - numA;
            }

            if (valA < valB) return window.currentSort.direction === 'asc' ? -1 : 1;
            if (valA > valB) return window.currentSort.direction === 'asc' ? 1 : -1;
            return 0;
        });
    }

    renderTable(filtered);
    updateSortIcons();

    // 필터 연동 업데이트 반영 (다른 조건에 의해 기존 선택값이 무효화되었으면 필터 재적용)
    const filtersReset = updateColumnFilterOptions();
    if (filtersReset) {
        filterData();
    }
}

function updateColumnFilterOptions() {
    let filtersReset = false;
    for (const key in window.columnFilters) {
        if (!window.columnFilters[key] || window.columnFilters[key].length === 0) continue;

        let dataForThisColumn = window.allProductionData || [];
        // (직 선택 필터가 제거되었으므로 해당 부분 로직 제거)
        for (const k in window.columnFilters) {
            if (k !== key && window.columnFilters[k] && window.columnFilters[k].length > 0) {
                dataForThisColumn = dataForThisColumn.filter(item => window.columnFilters[k].includes(String(item[k])));
            }
        }

        const uniqueValues = [...new Set(dataForThisColumn.map(item => {
            let itemVal = item[key];
            if (itemVal === undefined || itemVal === null) itemVal = '';
            return String(itemVal);
        }))];

        const originalSelection = window.columnFilters[key];
        const validSelection = originalSelection.filter(val => uniqueValues.includes(String(val)));

        if (validSelection.length !== originalSelection.length) {
            window.columnFilters[key] = validSelection;
            filtersReset = true;
        }
    }
    return filtersReset;
}

function updateSortIcons() {
    const headers = document.querySelectorAll('#production-table th');
    headers.forEach(th => {
        const key = th.getAttribute('data-key');

        // 화살표 제거 및 아이콘 상태 갱신
        let oldSortIcon = th.querySelector('.sort-icon');
        if (oldSortIcon) {
            oldSortIcon.remove();
        }

        // 현재 필터 적용 여부에 맞춰 아이콘 색상 변경
        let filterIcon = th.querySelector('.header-filter-icon');
        if (filterIcon) {
            if (window.columnFilters[key] && window.columnFilters[key].length > 0) {
                filterIcon.classList.add('active');
            } else {
                filterIcon.classList.remove('active');
            }
        }

        if (key && window.currentSort.key === key) {
            const icon = document.createElement('span');
            icon.className = 'sort-icon';
            if (window.currentSort.direction === 'asc') {
                icon.innerText = '▲';
                icon.style.color = '#10b981';
            } else {
                icon.innerText = '▼';
                icon.style.color = '#ef4444';
            }

            // 컨테이너가 있으면 그 안에, 필터 아이콘 앞에 삽입
            let actionsContainer = th.querySelector('.header-actions');
            if (actionsContainer) {
                if (filterIcon) {
                    actionsContainer.insertBefore(icon, filterIcon);
                } else {
                    actionsContainer.appendChild(icon);
                }
            } else {
                th.appendChild(icon);
            }
        }
    });
}

function updateClock() {
    const now = new Date();
    const options = { year: 'numeric', month: 'long', day: 'numeric', weekday: 'short' };
    const dateStr = now.toLocaleDateString('ko-KR', options);
    const weekNum = getWeekNumber(now);
    const monthWeekNum = getMonthWeekNumber(now);
    document.getElementById('current-date').innerText = `${dateStr} (${weekNum}주차 / ${now.getMonth() + 1}월 ${monthWeekNum}주차)`;

    // 주간 날짜 범위 업데이트 (월 ~ 금 기준)
    updateWeekRange(now);
}

function updateWeekRange(now) {
    const rangeElem = document.getElementById('week-range');
    if (!rangeElem) return;

    // 이번 주 월요일 찾기 (요일: 0=일, 1=월, ...)
    const day = now.getDay();
    const diffToMonday = (day === 0 ? -6 : 1 - day); // 일요일이면 -6, 그외 1-요일

    const monday = new Date(now);
    monday.setDate(now.getDate() + diffToMonday);

    const friday = new Date(monday);
    friday.setDate(monday.getDate() + 4); // 금요일은 월요일 + 4일

    const MondayStr = `${monday.getMonth() + 1}/${monday.getDate()}`;
    const FridayStr = `${friday.getMonth() + 1}/${friday.getDate()}`;

    rangeElem.innerText = `(${MondayStr} ~ ${FridayStr})`;
}

function getWeekNumber(d) {
    // 날짜 사본 생성
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    // 해당 주의 목요일로 이동 (ISO 8601 기준)
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    // 연도 시작일 계산
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    // 주차 계산
    const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return weekNo;
}

function getMonthWeekNumber(d) {
    const firstDay = new Date(d.getFullYear(), d.getMonth(), 1);
    const date = d.getDate();
    const day = firstDay.getDay(); // 1일의 요일 (0: 일요일)

    // 해당 월의 몇 번째 주인지 계산
    return Math.ceil((date + day) / 7);
}


function renderTable(data) {
    const tbody = document.getElementById('table-body');
    if (!tbody) return; // 테이블 요소가 없는 페이지(예: 작업일보)에서는 중단
    tbody.innerHTML = '';

    data.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.no}</td>
            <td>${item.month}</td>
            <td>${item.section}</td>
            <td>${item.model}</td>
            <td style="font-weight: 600;">${item.serial}</td>
            <td>${item.order}</td>
            <td>${item.customer}</td>
            <td>${renderDateBadge(item.first_shipment)}</td>
            <td style="font-weight: 600; color: var(--primary-color);">${renderDateBadge(item.target)}</td>
            <td>${renderDateBadge(item.base)}</td>
            <td>${renderDateBadge(item.first_start)}</td>
            <td>${renderDateBadge(item.revised_start)}</td>
            <td>${renderDateBadge(item.nc)}</td>
            <td><span class="status-badge ${item.status.includes('완료') ? 'status-completed' : 'status-ongoing'}">${item.status}</span></td>
            <td style="color: var(--danger); font-size: 0.8rem; text-align: left; min-width: 150px; white-space: normal;">${item.issue !== '-' && item.issue !== 'nan' ? item.issue : ''}</td>
        `;
        tbody.appendChild(row);
    });
}

function renderDateBadge(date) {
    if (date === '-' || date === 'NaT') return '<span style="color: #cbd5e1">-</span>';
    return `<span>${date}</span>`;
}

function renderDashboardSummaries(data) {
    // 1. 직별 완료율 계산 및 출력
    const sectionListDiv = document.getElementById('section-completion-list');
    if (sectionListDiv) {
        // 직별 데이터 매핑 (0AT1 -> 1직)
        const targetSectionsMap = {
            '0AT1': '1직',
            '0AT2': '2직',
            '0AT3': '3직',
            '0AT4': '4직',
            '0AT5': '5직'
        };
        // 화면에 출력할 직 순서
        const displaySections = ['1직', '2직', '3직', '4직', '5직'];
        const sectionsData = {};

        data.forEach(item => {
            // 소문자로 입력된 경우도 처리하기 위해 대문자로 변환
            let secRaw = (item.section || '').trim().toUpperCase();

            // 매핑 테이블에서 찾거나, 이미 '1직' 형태로 들어온 경우 처리
            let sec = targetSectionsMap[secRaw];
            if (!sec && displaySections.includes(secRaw)) {
                sec = secRaw;
            }

            if (!sec) return;

            if (!sectionsData[sec]) sectionsData[sec] = { total: 0, completed: 0 };
            sectionsData[sec].total++;
            if (item.status && item.status.includes('완료')) {
                sectionsData[sec].completed++;
            }
        });

        // 1직~5직 순서대로 정렬하여 출력
        sectionListDiv.innerHTML = displaySections.map(sec => {
            const stats = sectionsData[sec] || { total: 0, completed: 0 };
            const rate = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;
            return `
                <div style="margin-bottom: 0.8rem;">
                    <div style="display: flex; justify-content: space-between; font-size: 0.85rem; margin-bottom: 3px;">
                        <span style="font-weight: 600;">${sec}</span>
                        <span style="color: var(--secondary-color);">${rate}% (${stats.completed}/${stats.total})</span>
                    </div>
                    <div style="height: 8px; background: #e2e8f0; border-radius: 4px; overflow: hidden;">
                        <div style="width: ${rate}%; height: 100%; background: ${rate === 100 ? '#22c55e' : '#2563eb'}; transition: width 0.5s;"></div>
                    </div>
                </div>
            `;
        }).join('');
    }

    // 2. 출하 임박 공정 요약 요소의 내용 비움 처리
    const imminentDiv = document.getElementById('weekly-schedule');
    if (imminentDiv) {
        imminentDiv.innerHTML = ''; // 기존 목업 차트 제거
    }
}

// ==========================================
// 4. 컬럼 헤더 정렬 로직
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
    const table = document.getElementById('production-table');
    if (!table) return;

    const headers = table.querySelectorAll('th');

    headers.forEach(header => {
        header.style.cursor = 'pointer';
        header.addEventListener('click', handleHeaderClick);
    });
});

function handleHeaderClick(e) {
    // 팝업, 선택 상자, 혹은 그 내부 요소 클릭일 경우 정렬 무시
    if (e.target.closest('.header-filter-icon') || e.target.closest('.excel-filter-menu')) {
        return;
    }

    const key = this.getAttribute('data-key');
    if (!key) return;

    if (window.currentSort.key === key) {
        window.currentSort.direction = window.currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
        window.currentSort.key = key;
        window.currentSort.direction = 'asc';
    }

    filterData();
}

// ==========================================
// 5. 엑셀(CSV) 다운로드 기능
// ==========================================
function downloadExcel() {
    console.log("엑셀 다운로드 버튼 클릭됨");

    const table = document.getElementById('production-table');
    if (!table) {
        alert("다운로드할 데이터가 없습니다.");
        return;
    }

    // 한글 깨짐 방지를 위한 BOM (UTF-8)
    let csvContent = '\uFEFF';

    const rows = table.querySelectorAll('tr');
    if (rows.length === 0) {
        alert("다운로드할 데이터가 없습니다.");
        return;
    }

    rows.forEach(row => {
        const rowData = [];
        const cols = row.querySelectorAll('th, td');

        cols.forEach(col => {
            // 헤더에서 버튼/아이콘 제거 후 텍스트만 추출
            const clone = col.cloneNode(true);
            const filterIcon = clone.querySelector('.header-filter-icon');
            if (filterIcon) {
                clone.removeChild(filterIcon);
            }
            const sortIcon = clone.querySelector('.sort-icon');
            if (sortIcon) {
                clone.removeChild(sortIcon);
            }

            let text = clone.innerText.trim();

            // CSV 처리: 큰따옴표 이스케이프 및 쉼표/줄바꿈 포함 시 감싸기
            text = text.replace(/"/g, '""');
            if (text.includes(',') || text.includes('\n') || text.includes('"')) {
                text = '"' + text + '"';
            }
            rowData.push(text);
        });

        csvContent += rowData.join(',') + '\n';
    });

    // Blob 생성 및 다운로드 링크 처리
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');

    // 파일명 설정 (오늘 날짜)
    const today = new Date();
    const dateStr = today.getFullYear() +
        String(today.getMonth() + 1).padStart(2, '0') +
        String(today.getDate()).padStart(2, '0');

    link.setAttribute('href', url);
    link.setAttribute('download', `생산데이터_${dateStr}.csv`);

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    console.log("엑셀 다운로드 수행 완료");
}

// ==========================================
// 6. 사이드바 사이드 메뉴 토글 로직
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
    const sidebar = document.getElementById('mainSidebar');
    const toggleBtn = document.getElementById('sidebarToggle');
    const toggleIcon = document.getElementById('toggleIcon');

    if (toggleBtn && sidebar && toggleIcon) {
        toggleBtn.addEventListener('click', () => {
            sidebar.classList.toggle('collapsed');
        });
    }
});
