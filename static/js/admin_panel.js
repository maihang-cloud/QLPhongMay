
    // Cập nhật URL tải Báo cáo Tích hợp Tháng (Excel & Word) khi chọn tháng/năm
    function updateMonthlyExportLinks() {
        const mEl = document.getElementById('monthlyReportMonth');
        const yEl = document.getElementById('monthlyReportYear');
        const m = mEl ? mEl.value : 11;
        const y = yEl ? yEl.value : 2026;
        const btnExcel = document.getElementById('btnExportMonthlyExcel');
        const btnWord = document.getElementById('btnExportMonthlyWord');
        if (btnExcel) {
            btnExcel.href = `${window.ADMIN_PANEL_CFG.exportMonthlyExcelUrl}?month=${m}&year=${y}`;
        }
        if (btnWord) {
            btnWord.href = `${window.ADMIN_PANEL_CFG.exportMonthlyWordUrl}?month=${m}&year=${y}`;
        }
    }

    // ====================================================
    // BỘ LỌC, TÌM KIẾM & CHỌN XÓA HÀNG LOẠT (3 TAB ADMIN)
    // ====================================================
    const currentTabStatusFilter = {
        requests: 'all',
        issues: 'all',
        cancellations: 'all'
    };

    function filterAdminList(tabName, statusVal, btnEl) {
        currentTabStatusFilter[tabName] = statusVal;
        const buttons = document.querySelectorAll(`button[data-tab-filter="${tabName}"]`);
        buttons.forEach(b => {
            b.classList.remove('active-all', 'active-open', 'active-locked');
        });
        if (btnEl) {
            btnEl.classList.add('active-all');
        }
        applyAdminTabFilters(tabName);
    }

    function applyAdminTabFilters(tabName) {
        const statusFilter = currentTabStatusFilter[tabName] || 'all';
        const searchInput = document.getElementById(`searchInput-${tabName}`);
        const roomSelect = document.getElementById(`roomFilter-${tabName}`);
        const keyword = (searchInput ? searchInput.value : '').trim().toLowerCase();
        const roomVal = roomSelect ? roomSelect.value : 'all';

        const items = document.querySelectorAll(`.admin-list-item-${tabName}`);
        items.forEach(item => {
            const itemStatus = item.getAttribute('data-status') || 'Pending';
            const itemRoom = item.getAttribute('data-room') || '';
            const itemSearch = item.getAttribute('data-search') || '';

            const matchStatus = (statusFilter === 'all') || (itemStatus === statusFilter);
            const matchRoom = (roomVal === 'all') || (itemRoom === roomVal);
            const matchSearch = (!keyword) || itemSearch.includes(keyword);

            if (matchStatus && matchRoom && matchSearch) {
                item.style.display = '';
            } else {
                item.style.display = 'none';
            }
        });
        updateSelectedCount(tabName);
    }

    function syncRowCheckbox(tabName, changedCheckbox) {
        const val = changedCheckbox.value;
        const isChecked = changedCheckbox.checked;
        // Đồng bộ trạng thái tích giữa thẻ Mobile và dòng Desktop của cùng 1 ID
        document.querySelectorAll(`.item-check-${tabName}[value="${val}"]`).forEach(cb => {
            cb.checked = isChecked;
        });
        updateSelectedCount(tabName);
    }

    function toggleSelectAll(tabName, isChecked) {
        const items = document.querySelectorAll(`.admin-list-item-${tabName}`);
        const visibleIds = new Set();
        items.forEach(item => {
            if (item.style.display !== 'none') {
                const cb = item.querySelector(`.item-check-${tabName}`);
                if (cb) visibleIds.add(cb.value);
            }
        });

        document.querySelectorAll(`.item-check-${tabName}`).forEach(cb => {
            if (visibleIds.has(cb.value)) {
                cb.checked = isChecked;
            }
        });

        const desktopAll = document.getElementById(`selectAllDesktop-${tabName}`);
        const mobileAll = document.getElementById(`selectAllMobile-${tabName}`);
        if (desktopAll) desktopAll.checked = isChecked;
        if (mobileAll) mobileAll.checked = isChecked;

        updateSelectedCount(tabName);
    }

    function updateSelectedCount(tabName) {
        const selectedIds = new Set();
        document.querySelectorAll(`.item-check-${tabName}:checked`).forEach(cb => {
            selectedIds.add(cb.value);
        });

        const countSpan = document.getElementById(`selectedCount-${tabName}`);
        const bulkBtn = document.getElementById(`btnBulkDelete-${tabName}`);
        const hiddenContainer = document.getElementById(`bulkHiddenInputs-${tabName}`);

        if (countSpan) countSpan.textContent = selectedIds.size;
        if (bulkBtn) bulkBtn.disabled = (selectedIds.size === 0);
        if (hiddenContainer) {
            hiddenContainer.innerHTML = '';
            selectedIds.forEach(id => {
                const input = document.createElement('input');
                input.type = 'hidden';
                input.name = 'selected_ids';
                input.value = id;
                hiddenContainer.appendChild(input);
            });
        }
    }

    function confirmBulkDelete(tabName, labelName) {
        const selectedIds = new Set();
        document.querySelectorAll(`.item-check-${tabName}:checked`).forEach(cb => {
            selectedIds.add(cb.value);
        });
        if (selectedIds.size === 0) {
            alert('Vui lòng tích chọn ít nhất 1 mục để xóa!');
            return false;
        }
        return confirm(`⚠️ Bạn có chắc chắn muốn xóa vĩnh viễn ${selectedIds.size} ${labelName} đã chọn không?`);
    }

    function applyScheduleFilters() {
        const searchInput = document.getElementById('searchInput-schedules');
        const roomSelect = document.getElementById('roomFilter-schedules');
        const daySelect = document.getElementById('dayFilter-schedules');
        const clearBtn = document.getElementById('btnClearSearch-schedules');

        const keyword = (searchInput ? searchInput.value : '').trim().toLowerCase();
        const roomVal = roomSelect ? roomSelect.value : 'all';
        const dayVal = daySelect ? daySelect.value : 'all';

        if (clearBtn) {
            if (keyword || roomVal !== 'all' || dayVal !== 'all') {
                clearBtn.classList.remove('d-none');
            } else {
                clearBtn.classList.add('d-none');
            }
        }

        const rows = document.querySelectorAll('.admin-list-item-schedules');
        let visibleCount = 0;

        rows.forEach(row => {
            const rRoom = row.getAttribute('data-room') || '';
            const rDay = row.getAttribute('data-day') || '';
            const rSearch = row.getAttribute('data-search') || '';

            const matchRoom = (roomVal === 'all') || (rRoom === roomVal);
            const matchDay = (dayVal === 'all') || (rDay === dayVal);
            const matchKeyword = (!keyword) || rSearch.includes(keyword);

            if (matchRoom && matchDay && matchKeyword) {
                row.style.display = '';
                visibleCount++;
            } else {
                row.style.display = 'none';
                const cb = row.querySelector('.item-check-schedules');
                if (cb && cb.checked) cb.checked = false;
            }
        });

        const countEl = document.getElementById('visibleScheduleCount');
        if (countEl) countEl.textContent = visibleCount;

        const noMatchRow = document.getElementById('noScheduleMatchRow');
        if (noMatchRow && rows.length > 0) {
            if (visibleCount === 0) {
                noMatchRow.classList.remove('d-none');
            } else {
                noMatchRow.classList.add('d-none');
            }
        }

        const desktopAll = document.getElementById('selectAllDesktop-schedules');
        if (desktopAll) desktopAll.checked = false;
        updateSelectedCount('schedules');
    }

    function clearScheduleFilters() {
        const searchInput = document.getElementById('searchInput-schedules');
        const roomSelect = document.getElementById('roomFilter-schedules');
        const daySelect = document.getElementById('dayFilter-schedules');
        if (searchInput) searchInput.value = '';
        if (roomSelect) roomSelect.value = 'all';
        if (daySelect) daySelect.value = 'all';
        applyScheduleFilters();
    }

    function openUpdateWeeksModal(scheduleId, subjectName, roomLabel, currentWeeks, returnTab) {
        const form = document.getElementById('updateWeeksForm');
        form.action = `/admin/schedule/${scheduleId}/update_weeks`;
        document.getElementById('uw_course_label').textContent = `${subjectName} (${roomLabel})`;
        document.getElementById('uw_num_weeks').value = currentWeeks;
        document.getElementById('uw_return_tab').value = returnTab || 'schedules';
        const modal = new bootstrap.Modal(document.getElementById('updateWeeksModal'));
        modal.show();
    }

    function handleExportSemesterChange() {
        const selectEl = document.getElementById('exportSemesterSelect');
        if (!selectEl) return;
        const selectedOpt = selectEl.options[selectEl.selectedIndex];
        const val = selectEl.value;
        const count = parseInt(selectedOpt.getAttribute('data-count') || '0', 10);
        const yearVal = selectedOpt.getAttribute('data-year') || '';

        const customBox = document.getElementById('customDateRangeBox');
        const warnBox = document.getElementById('emptySemesterWarning');
        const submitBtn = document.getElementById('btnSubmitExportReport');
        const yearInput = document.getElementById('exportAcademicYearInput');

        if (yearInput && yearVal) {
            yearInput.value = yearVal;
        }

        if (customBox) {
            if (val === 'custom') {
                customBox.classList.remove('d-none');
            } else {
                customBox.classList.add('d-none');
            }
        }

        if (warnBox && submitBtn) {
            if (val !== 'custom' && count === 0) {
                warnBox.classList.remove('d-none');
                submitBtn.disabled = true;
            } else {
                warnBox.classList.add('d-none');
                submitBtn.disabled = false;
            }
        }
    }

    // =========================================================================
    // [#QTTB-03] CONTROLLER TRUNG TÂM GIÁM SÁT 8 ROUTER GRANDSTREAM (GCC6010 & GWN7062G)
    // =========================================================================
    window._latestRouterRoomMap = window._latestRouterRoomMap || {};
    window._currentDetailModalRoom = null;
    window._currentDeviceFilterMode = 'all';

    function escapeRouterHtml(str) {
        const div = document.createElement('div');
        div.textContent = str == null ? '' : String(str);
        return div.innerHTML;
    }

    function applyNetworkStatusData(data) {
        if (!data || !data.room_map) return;
        window._latestRouterRoomMap = data.room_map;
        const roomsList = ['5.1', '5.2', '5.3', '6.1', '6.2', '6.3', '6.5', '6.6'];

        roomsList.forEach(r => {
            const safeId = r.replace('.', '_');
            const nodeEl = document.getElementById(`room-node-${safeId}`);
            const pingEl = document.getElementById(`room-ping-${safeId}`);
            const modelEl = document.getElementById(`room-model-${safeId}`);
            const stateBadgeEl = document.getElementById(`room-state-badge-${safeId}`);
            const ipEl = document.getElementById(`room-ip-${safeId}`);
            const bwEl = document.getElementById(`room-bw-${safeId}`);
            const sysEl = document.getElementById(`room-sys-${safeId}`);

            const info = data.room_map[r];
            if (!info || !nodeEl || !pingEl) return;

            if (modelEl && info.router_model) {
                modelEl.textContent = info.router_model;
            }
            if (ipEl) {
                ipEl.innerHTML = `WAN: <strong class="text-dark">${escapeRouterHtml(info.target_ip)}:${info.https_port || 443}</strong> • LAN: <span class="text-secondary">${escapeRouterHtml(info.lan_ip || '-')}</span>`;
            }
            if (bwEl) {
                bwEl.innerHTML = `
                    <span class="fw-semibold" style="color: #0F766E;">⬇ ${escapeRouterHtml(info.rx_rate_str || '0 B/s')} | ⬆ ${escapeRouterHtml(info.tx_rate_str || '0 B/s')}</span>
                    <span class="text-muted" style="font-size: 0.67rem;">PC: <strong>${info.online_clients || 0}</strong> (ARP: ${info.arp_count || 0})</span>
                `;
            }
            if (sysEl) {
                const ls = info.live_session || {};
                if (ls.subject_name) {
                    const deptStr = ls.department ? ` (${ls.department})` : '';
                    sysEl.textContent = `${ls.ca_label || ''} • ${ls.subject_name}${deptStr} • ${ls.active_duration_str || '0 phút'}`;
                } else if (info.cpu_usage || info.mem_usage || (info.telemetry && info.telemetry.connected_api)) {
                    sysEl.textContent = `CPU ${info.cpu_usage || 0}% • RAM ${info.mem_usage || 0}% • ${info.temp_c || 38}°C`;
                } else if (ls.ca_label) {
                    sysEl.textContent = `${ls.ca_label} • ${ls.active_duration_str || '0 phút'}`;
                }
            }

            const rState = info.room_state || (info.status === 'online' ? 'wan_ping_only' : 'offline');
            if (stateBadgeEl) {
                let badgeCls = 'bg-info-subtle text-info-emphasis border border-info-subtle';
                if (rState === 'open_active') {
                    badgeCls = 'bg-success-subtle text-success border border-success-subtle';
                } else if (rState === 'closed_cb_off') {
                    badgeCls = 'bg-warning-subtle text-warning-emphasis border border-warning-subtle';
                } else if (rState === 'offline') {
                    badgeCls = 'bg-danger-subtle text-danger border border-danger-subtle';
                }
                stateBadgeEl.className = `badge w-100 text-start py-1 px-2 fw-semibold text-wrap lh-sm ${badgeCls}`;
                stateBadgeEl.textContent = `● ${info.room_state_label || (info.status === 'online' ? 'Online' : 'Offline')}`;
            }

            nodeEl.classList.remove('status-online', 'status-warn', 'status-offline');
            pingEl.classList.remove('ping-ok', 'ping-warn', 'ping-err', 'ping-wait');

            if (info.status === 'online') {
                const pingVal = info.ping_ms != null ? `${Math.round(info.ping_ms)}ms` : 'Online';
                if (rState === 'closed_cb_off' || (info.ping_ms != null && info.ping_ms > 150)) {
                    nodeEl.classList.add('status-warn');
                    pingEl.classList.add('ping-warn');
                    pingEl.textContent = pingVal;
                } else {
                    nodeEl.classList.add('status-online');
                    pingEl.classList.add('ping-ok');
                    pingEl.textContent = pingVal;
                }
            } else {
                nodeEl.classList.add('status-offline');
                pingEl.classList.add('ping-err');
                pingEl.textContent = 'Offline';
            }
        });

        const kpiVal = document.getElementById('kpiNetworkVal');
        const kpiSub = document.getElementById('kpiNetworkSub');
        const lastLbl = document.getElementById('lastNetUpdateLabel');

        if (lastLbl && data.last_update) {
            lastLbl.textContent = data.last_update;
        }
        if (kpiVal) {
            const hasOff = (data.offline_rooms && data.offline_rooms.length > 0);
            kpiVal.innerHTML = `${data.online_count}/${data.total_rooms || 8} <span class="fs-6 fw-semibold ${hasOff ? 'text-danger' : 'text-success'}">Online</span>`;
        }
        if (kpiSub) {
            if (data.offline_rooms && data.offline_rooms.length > 0) {
                kpiSub.innerHTML = `<span class="text-danger fw-bold">${data.offline_rooms.join(', ')} mất kết nối</span>`;
            } else {
                let openCnt = 0, idleCnt = 0, cbOffCnt = 0;
                Object.values(roomMap).forEach(ri => {
                    if (ri.room_state === 'open_active') openCnt += 1;
                    else if (ri.room_state === 'lan_up_idle') idleCnt += 1;
                    else if (ri.room_state === 'closed_cb_off') cbOffCnt += 1;
                });
                kpiSub.innerHTML = `<span class="text-success fw-semibold">${openCnt} mở học • ${idleCnt} bật Switch • ${cbOffCnt} tắt CB</span>`;
            }
        }

        if (window._currentDetailModalRoom) {
            renderRouterDetailModalContent(window._currentDetailModalRoom);
        }
    }

    function openRouterConfigModal() {
        const modalEl = document.getElementById('routerConfigModal');
        if (!modalEl) return;
        const bsModal = bootstrap.Modal.getOrCreateInstance(modalEl);
        bsModal.show();
    }

    async function openRouterDetailModal(room) {
        window._currentDetailModalRoom = room;
        window._currentDeviceFilterMode = 'all';
        const searchIn = document.getElementById('rdDeviceSearchInput');
        if (searchIn) searchIn.value = '';
        const otherToggle = document.getElementById('rdShowOtherDevicesToggle');
        if (otherToggle) otherToggle.checked = false;

        const modalEl = document.getElementById('routerDetailModal');
        if (modalEl) {
            const bsModal = bootstrap.Modal.getOrCreateInstance(modalEl);
            bsModal.show();
        }

        // Nếu chưa có telemetry chi tiết cho phòng này thì tự động quét ngay
        const existing = window._latestRouterRoomMap[room];
        if (!existing || !existing.telemetry || (!existing.telemetry.ports || existing.telemetry.ports.length === 0)) {
            renderRouterDetailModalContent(room, true);
            await scanSingleRouterFromModal(room);
        } else {
            renderRouterDetailModalContent(room, false);
        }
    }

    function renderRouterDetailModalContent(room, isLoading) {
        const info = window._latestRouterRoomMap[room] || { room: room, target_ip: '-', https_port: 443, router_model: 'Grandstream', lan_ip: '-', telemetry: {} };
        const telem = info.telemetry || {};

        const modelBadge = document.getElementById('rdModalModelBadge');
        const titleEl = document.getElementById('rdModalTitle');
        const subEl = document.getElementById('rdModalSub');
        const portsGrid = document.getElementById('rdPortsGrid');
        const diagBox = document.getElementById('rdDiagnosticBox');
        const metricsGrid = document.getElementById('rdMetricsGrid');

        const fwStr = telem.firmware && telem.firmware !== '-' ? ` • FW v${telem.firmware}` : '';
        if (modelBadge) modelBadge.textContent = `${info.router_model || 'Grandstream'}${fwStr}`;
        if (titleEl) titleEl.textContent = `CHI TIẾT ROUTER PHÒNG ${room}`;
        if (subEl) {
            subEl.textContent = `WAN: ${info.target_ip || '-'}:${info.https_port || 443}  |  LAN: ${info.lan_ip || '-'}/24  |  Trạng thái: ${info.room_state_label || 'Đang kiểm tra...'}`;
        }

        if (isLoading && portsGrid) {
            portsGrid.innerHTML = `<div class="text-muted small py-2"><span class="spinner-border spinner-border-sm me-2 text-primary"></span>Đang kết nối WebSocket tới Router Phòng ${room}...</div>`;
        } else if (portsGrid) {
            const ports = telem.ports || [];
            if (ports.length === 0) {
                portsGrid.innerHTML = `<div class="text-muted small py-2">Chưa có dữ liệu cổng vật lý (Cổng HTTPS ${info.https_port || 443} trên Router Phòng ${room} chưa mở Web WAN Access).</div>`;
            } else {
                portsGrid.innerHTML = ports.map(p => {
                    const bg = p.is_up ? '#DCFCE7' : '#E2E8F0';
                    const bd = p.is_up ? '#16A34A' : '#94A3B8';
                    const fg = p.is_up ? '#15803D' : '#475569';
                    const iconColor = p.is_up ? '#16A34A' : '#64748B';
                    return `
                        <div class="text-center p-2 rounded-3" style="min-width: 74px; background:${bg}; border: 1.5px solid ${bd};">
                            <div class="mx-auto mb-1 rounded-1" style="width: 32px; height: 14px; background:${iconColor};"></div>
                            <div class="fw-bold text-dark" style="font-size: 0.74rem;">${escapeRouterHtml(p.port_name)}</div>
                            <div class="fw-semibold" style="font-size: 0.65rem; color:${fg};">${escapeRouterHtml(p.role)}•${p.is_up ? 'UP' : 'DOWN'}</div>
                            <div style="font-size: 0.63rem; color:${fg};">${escapeRouterHtml(p.speed_label)}</div>
                        </div>
                    `;
                }).join('');
            }
        }

        if (diagBox) {
            diagBox.innerHTML = `<strong><i class="bi bi-info-circle-fill me-1"></i>Chẩn đoán tự động Phòng ${room}:</strong><br>${escapeRouterHtml(telem.diagnostic_note || 'Bấm "Làm mới phòng này" để đọc thông số trực tiếp từ Router.')}`;
        }

        if (metricsGrid) {
            const mCards = [
                { title: 'Thiết bị & Serial', val: telem.model || info.router_model || '-', sub: `SN: ${telem.serial || '-'} • FW: ${telem.firmware || '-'}` },
                { title: 'Địa chỉ MAC & Uptime', val: telem.mac || '-', sub: `Uptime: ${telem.uptime_str || '-'}` },
                { title: 'Hiệu năng Router', val: `CPU: ${info.cpu_usage || 0}% | RAM: ${info.mem_usage || 0}%`, sub: `Nhiệt độ: ${info.temp_c || '-'}°C • Ping: ${info.ping_ms || 0}ms` },
                { title: 'Tổng Lưu lượng WAN', val: `Down: ${telem.total_rx_str || info.total_rx_str || '0 B'}`, sub: `Up: ${telem.total_tx_str || info.total_tx_str || '0 B'} • Tốc độ: ⬇${info.rx_rate_str || '0 B/s'}` },
            ];
            metricsGrid.innerHTML = mCards.map(c => `
                <div class="col-6">
                    <div class="p-2 rounded-3 bg-light border h-100">
                        <div class="text-muted" style="font-size: 0.67rem;">${escapeRouterHtml(c.title)}</div>
                        <div class="fw-bold text-dark text-truncate" style="font-size: 0.78rem;">${escapeRouterHtml(c.val)}</div>
                        <div class="text-teal fw-medium text-truncate" style="font-size: 0.67rem; color:#0D9488;">${escapeRouterHtml(c.sub)}</div>
                    </div>
                </div>
            `).join('');
        }

        renderRouterModalDevices();
    }

    function setRouterDeviceFilter(mode) {
        window._currentDeviceFilterMode = mode;
        const btnAll = document.getElementById('rdFilterAllBtn');
        const btnOn = document.getElementById('rdFilterOnlineBtn');
        if (btnAll) btnAll.classList.toggle('active', mode === 'all');
        if (btnOn) btnOn.classList.toggle('active', mode === 'online');
        renderRouterModalDevices();
    }

    function renderRouterModalDevices() {
        const room = window._currentDetailModalRoom;
        if (!room) return;
        const info = window._latestRouterRoomMap[room] || {};
        const telem = info.telemetry || {};
        const allDevices = telem.devices || [];
        const showOther = Boolean((document.getElementById('rdShowOtherDevicesToggle') || {}).checked);
        const pcDevices = allDevices.filter(d => d.is_lab_pc !== false);
        const otherDevices = allDevices.filter(d => d.is_lab_pc === false);
        const baseList = showOther ? allDevices : pcDevices;
        const onlineDevices = baseList.filter(d => d.is_online);

        const cntAllEl = document.getElementById('rdCountAll');
        const cntOnEl = document.getElementById('rdCountOnline');
        const cntOtherEl = document.getElementById('rdOtherCount');
        if (cntAllEl) cntAllEl.textContent = String(baseList.length);
        if (cntOnEl) cntOnEl.textContent = String(onlineDevices.length);
        if (cntOtherEl) cntOtherEl.textContent = String(otherDevices.length);

        let list = window._currentDeviceFilterMode === 'online' ? onlineDevices : baseList;
        const kw = ((document.getElementById('rdDeviceSearchInput') || {}).value || '').trim().toLowerCase();
        if (kw) {
            list = list.filter(d =>
                String(d.computer_name || '').toLowerCase().includes(kw) ||
                String(d.ip || '').toLowerCase().includes(kw) ||
                String(d.mac || '').toLowerCase().includes(kw) ||
                String(d.device_category || '').toLowerCase().includes(kw)
            );
        }

        const tbody = document.getElementById('rdDevicesTableBody');
        if (!tbody) return;

        if (list.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="7" class="text-center text-muted py-4">
                        ${allDevices.length === 0
                            ? 'Chưa có thiết bị nào trong danh sách (hoặc Router phòng này chưa mở kết nối giám sát).'
                            : (baseList.length === 0
                                ? 'Không có PC phòng nào trong danh sách (bật nút gạt "Hiện thiết bị khác / Wi-Fi" để xem thiết bị ngoài).'
                                : 'Không có thiết bị nào khớp bộ lọc hiện tại.')}
                    </td>
                </tr>
            `;
            return;
        }

        tbody.innerHTML = list.map((d, idx) => {
            const stt = String(idx + 1).padStart(2, '0');
            const isOn = !!d.is_online;
            const isPc = (d.is_lab_pc !== false);
            const catLabel = d.device_category || (isPc ? 'PC Phòng' : 'Thiết bị khác');
            const catBadge = isPc
                ? `<span class="badge text-bg-success" style="font-size:0.63rem; background-color:#0D9488 !important;">${escapeRouterHtml(catLabel)}</span>`
                : `<span class="badge bg-secondary-subtle text-secondary border" style="font-size:0.63rem;">${escapeRouterHtml(catLabel)}</span>`;
            const statusDot = isOn
                ? `<span class="badge bg-success-subtle text-success border border-success-subtle me-1" style="font-size:0.63rem;">● Online</span>`
                : `<span class="badge bg-secondary-subtle text-secondary border me-1" style="font-size:0.63rem;">ARP Cache</span>`;
            const bwInfo = isOn
                ? `<span class="fw-semibold text-primary">⬇ ${escapeRouterHtml(d.down_rate_str)} | ⬆ ${escapeRouterHtml(d.up_rate_str)}</span><br><small class="text-muted">Tổng: ${escapeRouterHtml(d.total_bytes_str)}</small>`
                : `<span class="text-muted" style="font-size:0.69rem;">Lưu vết ca trước</span>`;

            return `
                <tr>
                    <td class="text-muted fw-semibold">${stt}</td>
                    <td>
                        <strong class="text-dark">${escapeRouterHtml(d.computer_name || '-')}</strong>
                        ${d.os ? `<br><small class="text-muted">${escapeRouterHtml(d.os)}</small>` : ''}
                    </td>
                    <td class="fw-bold text-primary font-monospace">${escapeRouterHtml(d.ip || '-')}</td>
                    <td class="font-monospace text-secondary" style="font-size:0.72rem;">${escapeRouterHtml(d.mac || '-')}</td>
                    <td>${catBadge}</td>
                    <td>
                        ${statusDot}
                        <span class="text-muted" style="font-size:0.68rem;">${escapeRouterHtml(d.vlan || 'VLAN 1')}</span>
                    </td>
                    <td class="text-end">${bwInfo}</td>
                </tr>
            `;
        }).join('');
    }

    async function savePcMacOuisFromModal() {
        const inp = document.getElementById('rdPcMacOuisInput');
        const btn = document.getElementById('rdBtnSaveOuis');
        if (!inp) return;
        if (btn) {
            btn.disabled = true;
            btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Đang lưu...';
        }
        try {
            const res = await fetch('/api/router-monitor/pc-ouis', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ pc_mac_ouis: inp.value })
            });
            const data = await res.json();
            if (data && data.success) {
                inp.value = data.pc_mac_ouis;
                await scanSingleRouterFromModal();
            }
        } catch (err) {
            console.error(err);
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = '<i class="bi bi-check2 me-1"></i>Lưu &amp; Quét lại';
            }
        }
    }

    async function scanSingleRouterFromModal(roomOverride) {
        const room = roomOverride || window._currentDetailModalRoom;
        if (!room) return;
        const btn = document.getElementById('rdBtnScanSingle');
        if (btn) {
            btn.disabled = true;
            btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Đang quét...';
        }
        try {
            await fetch('/api/router-monitor/scan-now', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ room: room })
            });
            await syncNetworkStatus();
        } catch (err) {
            console.error(err);
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = '<i class="bi bi-arrow-clockwise me-1"></i>Làm mới phòng này';
            }
        }
    }

    async function triggerRebootRouterFromModal() {
        const room = window._currentDetailModalRoom;
        if (!room) return;
        const info = window._latestRouterRoomMap[room] || {};
        if (!confirm(`Bạn có chắc chắn muốn KHỞI ĐỘNG LẠI (Reboot) Router ${info.router_model || ''} của Phòng ${room} (${info.target_ip || ''}) không?\nMạng phòng ${room} sẽ tạm ngắt khoảng 1-2 phút trong lúc Router khởi động lại.`)) {
            return;
        }
        const btn = document.getElementById('rdBtnReboot');
        if (btn) btn.disabled = true;
        try {
            const res = await fetch(`/api/router-monitor/reboot/${encodeURIComponent(room)}`, { method: 'POST' });
            const json = await res.json();
            alert((json && json.message) || 'Đã gửi yêu cầu.');
        } catch (err) {
            alert('Lỗi khi gửi lệnh Reboot: ' + err);
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    function updateBookingSwitchHint(el) {
        const hint = document.getElementById('bookingSwitchLabel');
        if (!hint) return;
        if (el.checked) {
            hint.innerHTML = '<span class="text-success fw-semibold"><i class="bi bi-check-circle-fill me-1"></i>Đang MỞ:</span> Giảng viên có thể truy cập trang và gửi đơn mượn phòng bình thường.';
        } else {
            hint.innerHTML = '<span class="text-danger fw-semibold"><i class="bi bi-lock-fill me-1"></i>Đang KHÓA:</span> Giảng viên sẽ thấy màn hình thông báo tạm ngưng và không thể gửi đơn.';
        }
    }

    async function syncNetworkStatus() {
        try {
            const res = await fetch('/api/network-status');
            const data = await res.json();
            if (data && data.success) {
                applyNetworkStatusData(data);
            }
        } catch (err) {
            console.warn('Network status sync error:', err);
        }
    }

    async function triggerPingNow() {
        const btn = document.getElementById('btnPingNow');
        if (btn) {
            btn.disabled = true;
            btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Đang quét 8 Router...';
        }
        try {
            await fetch('/api/router-monitor/scan-now', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({})
            });
            await syncNetworkStatus();
        } catch (err) {
            console.error(err);
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = '<i class="bi bi-arrow-repeat me-1"></i>Quét Realtime 8 Router';
            }
        }
    }

    // Activate tab based on URL parameter & start auto-sync
    document.addEventListener("DOMContentLoaded", function() {
        const urlParams = new URLSearchParams(window.location.search);
        const activeTab = urlParams.get('tab');
        if (activeTab) {
            const tabButton = document.querySelector(`button[data-bs-target="#${activeTab}"]`);
            if (tabButton) {
                const tab = new bootstrap.Tab(tabButton);
                tab.show();
            }
        }
        handleExportSemesterChange();
        
        // Auto-calculate OT hours from start and end time
        const sInput = document.getElementById('ot_start_time');
        const eInput = document.getElementById('ot_end_time');
        const hInput = document.getElementById('ot_hours');
        function calcOtHours() {
            if (sInput && eInput && hInput && sInput.value && eInput.value) {
                const [sh, sm] = sInput.value.split(':').map(Number);
                const [eh, em] = eInput.value.split(':').map(Number);
                let diff = (eh * 60 + em) - (sh * 60 + sm);
                if (diff < 0) diff += 24 * 60;
                const hours = Math.round((diff / 60) * 10) / 10;
                if (hours > 0) hInput.value = hours;
            }
        }
        if (sInput && eInput) {
            sInput.addEventListener('change', calcOtHours);
            eInput.addEventListener('change', calcOtHours);
        }

        // Lắng nghe sự kiện chuyển Tab để cập nhật URL và reset lại bộ đếm reload 30s
        const tabBtns = document.querySelectorAll('#adminTabs button[data-bs-toggle="tab"]');
        tabBtns.forEach(btn => {
            btn.addEventListener('shown.bs.tab', function(e) {
                const targetId = (e.target.getAttribute('data-bs-target') || '').replace('#', '');
                const url = new URL(window.location.href);
                url.searchParams.set('tab', targetId);
                history.replaceState(null, '', url.toString());
                
                autoReloadSeconds = 30;
                updateAutoReloadUI();
            });
        });

        startAutoReloadWatcher();
        setInterval(syncNetworkStatus, 15000);
        initDeployChecklistUI();
        applyContributorReadOnlyLock();
    });

    const IS_CONTRIBUTOR_READONLY = Boolean(window.ADMIN_PANEL_CFG.isContributorReadonly);

    function applyContributorReadOnlyLock() {
        if (!IS_CONTRIBUTOR_READONLY) return;

        // Vô hiệu hóa tất cả các form POST (thêm, sửa, duyệt, xóa, lưu cấu hình)
        document.querySelectorAll('form[method="POST"], form[method="post"]').forEach(form => {
            form.addEventListener('submit', function(e) {
                e.preventDefault();
                e.stopPropagation();
                alert('Tài khoản Contributor chỉ có quyền XEM, không có quyền chỉnh sửa, duyệt hoặc xóa!');
                return false;
            }, true);
            form.querySelectorAll('button, input, select, textarea').forEach(el => {
                el.disabled = true;
                el.title = 'Tài khoản Contributor chỉ có quyền xem';
            });
        });

        // Vô hiệu hóa các dropdown đổi trạng thái, checkbox chọn xóa và các nút mở Modal thêm/sửa/xóa
        const actionSelectors = [
            'select[onchange*="handleAdminStatusDropdownChange"]',
            'select[onchange*="handleScheduleDoorDropdownChange"]',
            'input[type="checkbox"][class*="item-check-"]',
            'input[type="checkbox"][id^="selectAll"]',
            'button[data-bs-target="#addScheduleModal"]',
            'button[data-bs-target="#importModal"]',
            'button[onclick*="openEditScheduleModal"]',
            'button[onclick*="openUpdateWeeksModal"]',
            'button[onclick*="openAddWorklogModal"]',
            'button[onclick*="openEditWorklogModal"]',
            'button[onclick*="completePhaseACurrentRoom"]',
            'button[onclick*="completeAllStepsCurrentRoom"]',
            'button[onclick*="resetCurrentRoomSteps"]',
            'button[onclick*="copyPhaseAToAllRooms"]',
            'button[onclick*="logCurrentDeployToWorklog"]',
            'button[onclick*="saveCurrentRoomDeployState"]',
            'button[onclick*="sendTestSmtpEmail"]',
            '#deployRoomNoteInput'
        ];
        document.querySelectorAll(actionSelectors.join(', ')).forEach(el => {
            el.disabled = true;
            el.style.opacity = '0.65';
            el.style.cursor = 'not-allowed';
            el.title = 'Tài khoản Contributor chỉ có quyền xem';
        });
    }

    // ==========================================
    // QUẢN LÝ NHẬT KÝ CÔNG VIỆC (WORK LOG) & CHECKLIST DEPLOY 8 PHÒNG MÁY
    // ==========================================
    let currentWorklogCategory = 'all';

    function switchWorklogSubView(mode) {
        const dailySec = document.getElementById('wlDailySection');
        const deploySec = document.getElementById('wlDeploySection');
        const btnDaily = document.getElementById('btnSubViewDaily');
        const btnDeploy = document.getElementById('btnSubViewDeploy');
        if (!dailySec || !deploySec) return;

        if (mode === 'deploy') {
            dailySec.classList.add('d-none');
            deploySec.classList.remove('d-none');
            btnDeploy.classList.add('bg-white', 'text-dark', 'shadow-sm');
            btnDeploy.classList.remove('text-white');
            btnDaily.classList.remove('bg-white', 'text-dark', 'shadow-sm');
            btnDaily.classList.add('text-white');
        } else {
            deploySec.classList.add('d-none');
            dailySec.classList.remove('d-none');
            btnDaily.classList.add('bg-white', 'text-dark', 'shadow-sm');
            btnDaily.classList.remove('text-white');
            btnDeploy.classList.remove('bg-white', 'text-dark', 'shadow-sm');
            btnDeploy.classList.add('text-white');
        }

        const url = new URL(window.location.href);
        url.searchParams.set('tab', 'worklog');
        url.searchParams.set('wl_view', mode);
        history.replaceState(null, '', url.toString());
    }

    function filterWorklogCategory(catVal, btnEl) {
        currentWorklogCategory = catVal;
        document.querySelectorAll('button[data-wl-cat]').forEach(b => {
            b.classList.remove('active-all');
        });
        if (btnEl) btnEl.classList.add('active-all');
        applyWorklogFilters();
    }

    function applyWorklogFilters() {
        const searchInput = document.getElementById('searchInput-worklogs');
        const roomSelect = document.getElementById('roomFilter-worklogs');
        const statusSelect = document.getElementById('statusFilter-worklogs');

        const keyword = (searchInput ? searchInput.value : '').trim().toLowerCase();
        const roomVal = roomSelect ? roomSelect.value : 'all';
        const statusVal = statusSelect ? statusSelect.value : 'all';

        const items = document.querySelectorAll('.admin-list-item-worklogs');
        items.forEach(item => {
            const itemCat = item.getAttribute('data-category') || '';
            const itemStatus = item.getAttribute('data-status') || '';
            const itemRoom = item.getAttribute('data-room') || '';
            const itemSearch = item.getAttribute('data-search') || '';

            const matchCat = (currentWorklogCategory === 'all') || (itemCat === currentWorklogCategory);
            const matchStatus = (statusVal === 'all') || (itemStatus === statusVal);
            const matchRoom = (roomVal === 'all') || itemRoom.toLowerCase().includes(roomVal.toLowerCase()) || itemRoom.includes('Tất cả');
            const matchSearch = (!keyword) || itemSearch.includes(keyword);

            if (matchCat && matchStatus && matchRoom && matchSearch) {
                item.style.display = '';
            } else {
                item.style.display = 'none';
            }
        });
        updateSelectedCount('worklogs');
    }

    function syncRoomChipsWithInput() {
        const inputEl = document.getElementById('wl_target_rooms');
        if (!inputEl) return;
        const currentParts = inputEl.value.split(',').map(s => s.trim()).filter(Boolean);
        document.querySelectorAll('#wlRoomChipsContainer .wl-room-chip').forEach(chip => {
            const val = chip.getAttribute('data-room-val');
            if (currentParts.includes(val)) {
                chip.classList.add('active');
            } else {
                chip.classList.remove('active');
            }
        });
    }

    function toggleWlRoomChip(chipEl) {
        chipEl.classList.toggle('active');
        const selected = [];
        document.querySelectorAll('#wlRoomChipsContainer .wl-room-chip.active').forEach(c => {
            selected.push(c.getAttribute('data-room-val'));
        });
        const inputEl = document.getElementById('wl_target_rooms');
        if (inputEl) {
            inputEl.value = selected.join(', ');
        }
    }

    function openAddWorklogModal() {
        const form = document.getElementById('worklogForm');
        form.action = window.ADMIN_PANEL_CFG.addWorklogUrl;
        document.getElementById('worklogModalTitle').textContent = 'Thêm Công Việc Thường Nhật Mới';
        document.getElementById('worklogSubmitBtn').textContent = 'Lưu vào Nhật ký công việc';

        document.getElementById('wl_work_date').value = window.ADMIN_PANEL_CFG.todayDateStr;
        document.getElementById('wl_shift_name').value = 'Ca sáng';
        document.getElementById('wl_priority').value = 'normal';
        document.getElementById('wl_category').value = 'server_net';
        document.getElementById('wl_status').value = 'in_progress';
        document.getElementById('wl_target_rooms').value = '5.1';
        document.getElementById('wl_content').value = '';
        document.getElementById('wl_notes').value = '';

        syncRoomChipsWithInput();
        const modal = new bootstrap.Modal(document.getElementById('worklogModal'));
        modal.show();
    }

    function openEditWorklogModal(btnEl) {
        const id = btnEl.getAttribute('data-id');
        const form = document.getElementById('worklogForm');
        form.action = `/admin/worklog/edit/${id}`;
        document.getElementById('worklogModalTitle').textContent = `Chỉnh Sửa Công Việc #${id}`;
        document.getElementById('worklogSubmitBtn').textContent = 'Lưu Cập Nhật';

        document.getElementById('wl_work_date').value = btnEl.getAttribute('data-date') || window.ADMIN_PANEL_CFG.todayDateStr;
        document.getElementById('wl_shift_name').value = btnEl.getAttribute('data-shift') || 'Ca sáng';
        document.getElementById('wl_priority').value = btnEl.getAttribute('data-priority') || 'normal';
        document.getElementById('wl_category').value = btnEl.getAttribute('data-category') || 'server_net';
        document.getElementById('wl_status').value = btnEl.getAttribute('data-status') || 'in_progress';
        document.getElementById('wl_target_rooms').value = btnEl.getAttribute('data-rooms') || '';
        document.getElementById('wl_content').value = btnEl.getAttribute('data-content') || '';
        document.getElementById('wl_notes').value = btnEl.getAttribute('data-notes') || '';

        syncRoomChipsWithInput();
        const modal = new bootstrap.Modal(document.getElementById('worklogModal'));
        modal.show();
    }

    // ====================================================
    // LOGIC CHECKLIST DEPLOY 8 PHÒNG MÁY (23 BƯỚC CHUẨN SOP)
    // ====================================================
    const DEPLOY_PHASE_A_STEPS = [
        { id: 'A0', label: 'A. Recovery từ file Acronis lên 2 máy mẫu (Sinh viên và Giáo viên)', hint: 'Lưu ý: File có sẵn Veyon dành cho Client' },
        { id: 'A1', label: '1. Kiểm tra Power Option', hint: 'Chỉnh High Performance, Turn off display / Sleep = Never' },
        { id: 'A2', label: '2. Cài đặt Driver', hint: 'Kiểm tra Device Manager nhận đủ toàn bộ thiết bị' },
        { id: 'A3', label: '3. Tắt Fast Startup', hint: 'Control Panel ➔ Power Options ➔ Choose what the power buttons do ➔ Bỏ tích Fast Startup' },
        { id: 'A4', label: '4. Kiểm tra Date and Time', hint: 'Đúng múi giờ (UTC+07:00) Bangkok, Hanoi, Jakarta & đồng bộ giờ chuẩn' },
        { id: 'A5', label: '5. Clean ổ D và This PC', hint: 'Dọn sạch dữ liệu cũ trên ổ D, Desktop, Downloads, Documents, Recycle Bin' },
        { id: 'A6', label: '6. Xóa Temp', hint: 'Chạy Run: temp, %temp%, prefetch và xóa sạch' },
        { id: 'A7', label: '7. Kiểm tra Windows Active', hint: 'Đảm bảo Windows & Office đã báo Activated bản quyền' },
        { id: 'A8', label: '8. Backup lại 1 file với đầy đủ các bước trên máy thực', hint: 'Để sau này cập nhật phần mềm chỉ cần cài phần mềm là okay' },
        { id: 'A9', label: '9. Chuyển sang Administrator ➔ Paste source Script Automation ➔ Run file "first run" ➔ PC Shutdown', hint: 'Thực hiện đúng trình tự trên tài khoản Administrator trước khi bắt Image' },
        { id: 'A10', label: '10. Tạo file Image chuẩn (7 máy giáo viên (Dell) và các máy sinh viên)', hint: 'Hoàn tất file Master Image chuẩn bị bung hàng loạt' }
    ];

    const DEPLOY_PHASE_B_STEPS = [
        { id: 'B1', label: '1. Kiểm tra tốc độ mạng của phòng phải đủ 1 Gbps', hint: 'Kiểm tra Link Speed = 1.0 Gbps trên máy mẫu & Switch trước khi chạy Clonezilla' },
        { id: 'B2', label: '2. Bung Image lên máy Giáo viên (GV) trước', hint: 'Khôi phục máy GV hoàn tất để dùng máy GV làm Clonezilla Lite Server' },
        { divider: '3. Từ máy giáo viên triển khai Clonezilla Lite Server lên máy con:' },
        { id: 'B3_1', sub: true, label: '3.1. Chọn mode: <span class="deploy-code-pill">lite-server</span>' },
        { id: 'B3_2', sub: true, label: '3.2. Chọn <span class="deploy-code-pill">Netboot</span>' },
        { id: 'B3_3', sub: true, label: '3.3. Chọn <span class="deploy-code-pill">Using existing DHCP</span>' },
        { id: 'B3_4', sub: true, label: '3.4. Chọn <span class="deploy-code-pill">DHCP use DHCP broadcast</span>' },
        { id: 'B3_5', sub: true, label: '3.5. Chọn <span class="deploy-code-pill">Choose number client</span> (Nhập đúng số máy con)' },
        { id: 'B3_6', sub: true, label: '3.6. Chọn <span class="deploy-code-pill">Reboot client</span> sau khi triển khai thành công' },
        { divider: '4. Cấu hình máy Giáo viên (GV):' },
        { id: 'B4_1', sub: true, label: '4.1. Đặt IP tĩnh cho máy Giáo viên' },
        { id: 'B4_2', sub: true, label: '4.2. Chép <span class="deploy-code-pill">t3h</span> và cài đặt phần mềm thi trắc nghiệm' },
        { id: 'B4_3', sub: true, label: '4.3. Chỉnh sửa quyền của ổ <span class="deploy-code-pill">Z</span> trong Logon, pass là <span class="deploy-secret-pill">giaovien</span>' },
        { id: 'B5', label: '5. Hoàn tất triển khai phòng máy', hint: 'Kiểm tra Veyon Master nhìn thấy đủ máy con & nghiệm thu phòng máy' }
    ];

    const ALL_DEPLOY_STEP_IDS = [
        ...DEPLOY_PHASE_A_STEPS.map(s => s.id),
        ...DEPLOY_PHASE_B_STEPS.filter(s => !s.divider).map(s => s.id)
    ];

    const DEPLOY_ROOMS_LIST = ['5.1', '5.2', '5.3', '6.1', '6.2', '6.3', '6.5', '6.6'];
    let deployRoomStateMap = window.ADMIN_PANEL_CFG.deployRoomStateMap;
    let activeDeployRoom = '5.1';

    function initDeployChecklistUI() {
        DEPLOY_ROOMS_LIST.forEach(r => {
            if (!deployRoomStateMap[r]) {
                deployRoomStateMap[r] = { room: r, checked: [], notes: '', updated_by: 'Admin', updated_at: '' };
            }
        });
        renderDeployUI();
    }

    function changeDeployCampaign(campaignName) {
        const url = new URL(window.location.href);
        url.searchParams.set('tab', 'worklog');
        url.searchParams.set('wl_view', 'deploy');
        url.searchParams.set('deploy_campaign', campaignName);
        window.location.href = url.toString();
    }

    function selectDeployRoom(roomCode) {
        activeDeployRoom = roomCode;
        renderDeployUI();
    }

    function renderDeployUI() {
        const gridEl = document.getElementById('deployRoomsGrid');
        if (!gridEl) return;

        const totalSteps = ALL_DEPLOY_STEP_IDS.length;
        let completedRoomsCount = 0;

        gridEl.innerHTML = DEPLOY_ROOMS_LIST.map(r => {
            const st = deployRoomStateMap[r] || { checked: [], notes: '' };
            const doneCount = (st.checked || []).length;
            const pct = Math.round((doneCount / totalSteps) * 100);
            const isDone = doneCount >= totalSteps;
            if (isDone) completedRoomsCount++;
            const isActive = (r === activeDeployRoom);

            return `
                <div class="deploy-room-card ${isActive ? 'active' : ''} ${isDone ? 'done' : ''}" onclick="selectDeployRoom('${r}')">
                    <div class="fw-bold" style="color: #0F2942; font-size: 0.86rem;">
                        Phòng ${r}
                    </div>
                    <div class="fw-bold mt-1" style="font-size: 0.74rem; color: ${isDone ? '#15803D' : (isActive ? '#0D9488' : '#64748B')};">
                        ${doneCount}/${totalSteps} bước (${pct}%)
                    </div>
                    <div class="progress mt-1" style="height: 5px; background-color: #E2E8F0;">
                        <div class="progress-bar" role="progressbar"
                             style="width: ${pct}%; background-color: ${isDone ? '#22C55E' : '#0D9488'};"></div>
                    </div>
                </div>
            `;
        }).join('');

        const summaryBadge = document.getElementById('overallDeploySummaryBadge');
        if (summaryBadge) {
            summaryBadge.textContent = `Tổng tiến độ đợt này: ${completedRoomsCount} / 8 phòng hoàn tất 100%`;
        }

        const curState = deployRoomStateMap[activeDeployRoom] || { checked: [], notes: '', updated_at: '' };
        const checkedSet = new Set(curState.checked || []);
        const curDone = checkedSet.size;
        const curPct = Math.round((curDone / totalSteps) * 100);

        document.getElementById('deployActiveRoomBadge').textContent = `PHÒNG ${activeDeployRoom}`;
        document.getElementById('deployNoteRoomLabel').textContent = `Phòng ${activeDeployRoom}`;
        document.getElementById('deployActiveRoomProgressText').textContent = `Tiến độ Phòng ${activeDeployRoom}: ${curDone} / ${totalSteps} bước (${curPct}%)`;
        document.getElementById('deployActiveRoomMeta').textContent = curState.updated_at
            ? `Cập nhật lần cuối: ${curState.updated_at} bởi ${curState.updated_by || 'Admin'}`
            : 'Tự động lưu vào cơ sở dữ liệu ngay khi tích chọn';
        document.getElementById('deployRoomNoteInput').value = curState.notes || '';

        // Render Phase A
        const phaseAHtml = DEPLOY_PHASE_A_STEPS.map(s => {
            const isChecked = checkedSet.has(s.id);
            return `
                <div class="deploy-step-item ${isChecked ? 'checked' : ''}" onclick="toggleDeployStep('${s.id}')">
                    <input type="checkbox" class="form-check-input mt-1 flex-shrink-0" ${isChecked ? 'checked' : ''}
                           onclick="event.stopPropagation(); toggleDeployStep('${s.id}')">
                    <div>
                        <div class="deploy-step-title fw-semibold text-dark" style="font-size: 0.82rem;">${s.label}</div>
                        ${s.hint ? `<div class="text-muted" style="font-size: 0.73rem;">${s.hint}</div>` : ''}
                    </div>
                </div>
            `;
        }).join('');
        document.getElementById('deployPhaseAList').innerHTML = phaseAHtml;

        // Render Phase B
        const phaseBHtml = DEPLOY_PHASE_B_STEPS.map(s => {
            if (s.divider) {
                return `<div class="fw-bold text-uppercase mt-2 mb-1 px-1" style="font-size: 0.73rem; color: #0F2942;">${s.divider}</div>`;
            }
            const isChecked = checkedSet.has(s.id);
            return `
                <div class="deploy-step-item ${s.sub ? 'sub-step' : ''} ${isChecked ? 'checked' : ''}" onclick="toggleDeployStep('${s.id}')">
                    <input type="checkbox" class="form-check-input mt-1 flex-shrink-0" ${isChecked ? 'checked' : ''}
                           onclick="event.stopPropagation(); toggleDeployStep('${s.id}')">
                    <div>
                        <div class="deploy-step-title fw-semibold text-dark" style="font-size: 0.82rem;">${s.label}</div>
                        ${s.hint ? `<div class="text-muted" style="font-size: 0.73rem;">${s.hint}</div>` : ''}
                    </div>
                </div>
            `;
        }).join('');
        document.getElementById('deployPhaseBList').innerHTML = phaseBHtml;

        const aDone = DEPLOY_PHASE_A_STEPS.filter(s => checkedSet.has(s.id)).length;
        const bCheckable = DEPLOY_PHASE_B_STEPS.filter(s => !s.divider);
        const bDone = bCheckable.filter(s => checkedSet.has(s.id)).length;

        document.getElementById('deployPhaseACountBadge').textContent = `${aDone}/${DEPLOY_PHASE_A_STEPS.length} bước`;
        document.getElementById('deployPhaseBCountBadge').textContent = `${bDone}/${bCheckable.length} bước`;
    }

    function toggleDeployStep(stepId) {
        if (IS_CONTRIBUTOR_READONLY) {
            alert('Tài khoản Contributor chỉ có quyền XEM, không có quyền tích chọn hoặc chỉnh sửa Checklist Deploy!');
            return;
        }
        const st = deployRoomStateMap[activeDeployRoom];
        const arr = st.checked || [];
        const idx = arr.indexOf(stepId);
        if (idx >= 0) {
            arr.splice(idx, 1);
        } else {
            arr.push(stepId);
        }
        st.checked = arr;
        renderDeployUI();
        saveCurrentRoomDeployState();
    }

    function completePhaseACurrentRoom() {
        const st = deployRoomStateMap[activeDeployRoom];
        const set = new Set(st.checked || []);
        DEPLOY_PHASE_A_STEPS.forEach(s => set.add(s.id));
        st.checked = Array.from(set);
        renderDeployUI();
        saveCurrentRoomDeployState();
    }

    function completeAllStepsCurrentRoom() {
        deployRoomStateMap[activeDeployRoom].checked = [...ALL_DEPLOY_STEP_IDS];
        renderDeployUI();
        saveCurrentRoomDeployState();
    }

    function resetCurrentRoomSteps() {
        if (!confirm(`Bạn có chắc muốn Reset toàn bộ tích chọn của Phòng ${activeDeployRoom} về 0% không?`)) return;
        deployRoomStateMap[activeDeployRoom].checked = [];
        renderDeployUI();
        saveCurrentRoomDeployState();
    }

    async function saveCurrentRoomDeployState(showNotice = false) {
        const campaignSelect = document.getElementById('deployCampaignSelect');
        const campaignName = campaignSelect ? campaignSelect.value : '';
        const noteInput = document.getElementById('deployRoomNoteInput');
        const st = deployRoomStateMap[activeDeployRoom];
        if (noteInput) {
            st.notes = noteInput.value;
        }

        try {
            const res = await fetch('/api/deploy-checklist/update', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    campaign_name: campaignName,
                    room: activeDeployRoom,
                    checked_steps: st.checked || [],
                    notes: st.notes || ''
                })
            });
            const data = await res.json();
            if (data && data.success) {
                st.updated_at = data.updated_at;
                st.updated_by = data.updated_by;
                document.getElementById('deployActiveRoomMeta').textContent =
                    `Đã lưu tự động lúc ${data.updated_at} bởi ${data.updated_by}`;
                if (showNotice) {
                    alert(`Đã lưu ghi chú và tiến độ của Phòng ${activeDeployRoom}!`);
                }
            }
        } catch (err) {
            console.error('Save deploy checklist error:', err);
        }
    }

    async function copyPhaseAToAll8Rooms() {
        if (!confirm('Đánh dấu hoàn thành 11 bước Phần A (Đã chuẩn bị xong Master Image) cho cả 8 phòng máy trong đợt này?')) return;
        const campaignSelect = document.getElementById('deployCampaignSelect');
        const campaignName = campaignSelect ? campaignSelect.value : '';
        try {
            const res = await fetch('/api/deploy-checklist/copy-phase-a-all', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ campaign_name: campaignName })
            });
            const data = await res.json();
            if (data && data.success && data.room_map) {
                deployRoomStateMap = data.room_map;
                renderDeployUI();
                alert('Đã đánh dấu xong Phần A (Master Image) cho toàn bộ 8 phòng máy!');
            }
        } catch (err) {
            console.error(err);
        }
    }

    async function logCurrentRoomDeployToWorklog() {
        const campaignSelect = document.getElementById('deployCampaignSelect');
        const campaignName = campaignSelect ? campaignSelect.value : '';
        const st = deployRoomStateMap[activeDeployRoom] || { checked: [], notes: '' };
        const noteVal = document.getElementById('deployRoomNoteInput').value || st.notes || '';
        const doneCount = (st.checked || []).length;
        const totalCount = ALL_DEPLOY_STEP_IDS.length;

        try {
            await saveCurrentRoomDeployState(false);
            const res = await fetch('/api/deploy-checklist/log-to-worklog', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    campaign_name: campaignName,
                    room: activeDeployRoom,
                    done_count: doneCount,
                    total_count: totalCount,
                    notes: noteVal
                })
            });
            const data = await res.json();
            if (data && data.success) {
                if (confirm(`${data.message}\n\nBạn có muốn chuyển sang xem Bảng Nhật ký công việc ngay không?`)) {
                    window.location.href = `/admin/panel?tab=worklog&wl_view=daily`;
                }
            }
        } catch (err) {
            console.error(err);
        }
    }

    // ==========================================
    // XỬ LÝ MODAL PHẢN HỒI & TRẢ KẾT QUẢ QUA EMAIL CHO USER
    // ==========================================
    let currentReplyContext = {};

    const QUICK_REPLY_TEMPLATES = {
        booking: {
            Processing: [
                'QTTB đã tiếp nhận yêu cầu và đang kiểm tra lịch phòng máy.',
                'Đang tiến hành cài đặt phần mềm theo yêu cầu của Quý Thầy/Cô.',
                'Đã tiếp nhận, sẽ phản hồi kết quả duyệt phòng sớm nhất.'
            ],
            Approved: [
                'Đã duyệt phòng và cập nhật vào Thời khóa biểu. Đã báo bộ phận trực mở cửa phòng đúng giờ.',
                'Đã duyệt mượn phòng và hoàn tất cài đặt phần mềm phục vụ giảng dạy.',
                'Phòng máy đã sẵn sàng, kính mời Quý Thầy/Cô sử dụng theo lịch đăng ký.'
            ],
            Rejected: [
                'Khung giờ này phòng máy đã có lịch học chính khóa hoặc lịch bảo trì định kỳ.',
                'Hiện tại tất cả các phòng máy đều đã kín lịch trong ca này, mong Quý Thầy/Cô thông cảm chọn khung giờ khác.',
                'Phần mềm yêu cầu không tương thích với cấu hình hiện tại của phòng máy này.'
            ]
        },
        issue: {
            Processing: [
                'Kỹ thuật viên QTTB đã tiếp nhận báo cáo và đang xuống phòng máy kiểm tra.',
                'Đã ghi nhận sự cố máy tính, kỹ thuật sẽ xử lý trước ca học tiếp theo.'
            ],
            Resolved: [
                'Đã kiểm tra và khắc phục hoàn tất sự cố phần cứng/phần mềm trên máy.',
                'Đã bung lại Image / cài đặt lại phần mềm, máy hoạt động ổn định bình thường.',
                'Đã thay thế thiết bị ngoại vi (chuột/bàn phím/dây mạng) mới cho máy.'
            ]
        },
        cancellation: {
            Processing: [
                'QTTB đã tiếp nhận thông báo nghỉ học và đang chuyển thông tin cho bộ phận trực phòng.',
                'Đã ghi nhận lịch báo vắng của Quý Thầy/Cô.'
            ],
            Approved: [
                'Đã xác nhận lịch báo nghỉ và thông báo chị Hân khóa cửa phòng máy trong ca này.',
                'Đã duyệt thông báo nghỉ và giải phóng ca phòng trống trên hệ thống. Quý Thầy/Cô có thể đăng ký dạy bù ở mục Mượn phòng.'
            ]
        }
    };

    const STATUS_DROPDOWN_LABELS = {
        booking: {
            Processing: 'Đã tiếp nhận yêu cầu mượn phòng',
            Approved: 'Đã duyệt yêu cầu mượn phòng',
            Rejected: 'Từ chối yêu cầu mượn phòng'
        },
        issue: {
            Processing: 'Đã tiếp nhận báo cáo sự cố',
            Resolved: 'Đã xong / Đã khắc phục sự cố'
        },
        cancellation: {
            Processing: 'Đã tiếp nhận thông báo nghỉ gấp',
            Approved: 'Đã duyệt thông báo nghỉ gấp'
        }
    };

    function handleScheduleDoorDropdownChange(selectEl) {
        const curVal = selectEl.getAttribute('data-current') || 'open';
        const newVal = selectEl.value;
        const subject = selectEl.getAttribute('data-subject') || 'Môn học';
        const room = selectEl.getAttribute('data-room') || '';

        if (newVal === 'locked') {
            if (!confirm(`Đánh dấu môn "${subject}" (P.${room}) kết thúc sớm? Hệ thống sẽ báo Khóa cửa phòng và giải phóng phòng trống.`)) {
                selectEl.value = curVal;
                return;
            }
        }
        selectEl.form.submit();
    }

    // ==========================================
    // XỬ LÝ THÊM LỊCH (1 MÔN HOẶC NHIỀU MÔN) & SỬA LỊCH MÔN HỌC
    // ==========================================
    let adminScheduleCart = [];

    function prepareAddScheduleModal(returnTab, defaultRoom) {
        const retEl = document.getElementById('addScheduleReturnTab');
        if (retEl) retEl.value = returnTab || 'schedules';
        if (defaultRoom) {
            const roomSel = document.getElementById('add_room');
            if (roomSel) roomSel.value = defaultRoom;
        }
    }

    function setAddScheduleInputsRequired(isReq) {
        ['add_start_period', 'add_end_period', 'add_subject_name', 'add_department'].forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                if (isReq) el.setAttribute('required', 'required');
                else el.removeAttribute('required');
            }
        });
    }

    function addScheduleToAdminCart() {
        const room = document.getElementById('add_room').value.trim();
        const day_of_week = document.getElementById('add_day_of_week').value;
        const start_period = parseFloat(document.getElementById('add_start_period').value);
        const end_period = parseFloat(document.getElementById('add_end_period').value);
        const subject_name = document.getElementById('add_subject_name').value.trim();
        const class_code = document.getElementById('add_class_code').value.trim();
        const department = document.getElementById('add_department').value;
        const study_period = document.getElementById('add_study_period').value.trim();
        const num_weeks = parseInt(document.getElementById('add_num_weeks').value) || 15;
        const notes = document.getElementById('add_notes').value.trim();

        if (!room || isNaN(start_period) || isNaN(end_period) || !subject_name || !department) {
            alert('Vui lòng nhập đầy đủ Phòng, Tiết BĐ, Tiết KT, Tên Môn và chọn Đơn vị trước khi thêm vào danh sách!');
            return;
        }
        if (start_period >= end_period) {
            alert('Tiết bắt đầu phải nhỏ hơn tiết kết thúc!');
            return;
        }
        if (start_period < 1 || end_period > 10) {
            alert('Tiết học phải nằm trong khoảng từ 1 đến 10!');
            return;
        }

        adminScheduleCart.push({
            room,
            day_of_week: parseInt(day_of_week),
            start_period,
            end_period,
            subject_name,
            class_code,
            department,
            study_period,
            num_weeks,
            notes
        });

        // Reset nhanh Mã lớp & Ghi chú để nhập ca tiếp theo nhanh hơn (giữ lại Tên môn, Đơn vị, Giai đoạn)
        document.getElementById('add_class_code').value = '';
        document.getElementById('add_notes').value = '';
        renderAdminScheduleCart();
    }

    function removeScheduleFromAdminCart(idx) {
        adminScheduleCart.splice(idx, 1);
        renderAdminScheduleCart();
    }

    function clearAdminScheduleCart() {
        adminScheduleCart = [];
        renderAdminScheduleCart();
    }

    function renderAdminScheduleCart() {
        const sec = document.getElementById('adminScheduleCartSection');
        const tbody = document.getElementById('adminScheduleCartBody');
        const countEl = document.getElementById('adminScheduleCartCount');
        const submitBtn = document.getElementById('btnSubmitAddSchedule');
        const bulkInput = document.getElementById('addScheduleBulkJson');

        if (!sec || !tbody) return;
        countEl.textContent = adminScheduleCart.length;

        if (adminScheduleCart.length === 0) {
            sec.classList.add('d-none');
            tbody.innerHTML = '';
            bulkInput.value = '';
            submitBtn.textContent = 'Lưu Lịch';
            setAddScheduleInputsRequired(true);
            return;
        }

        sec.classList.remove('d-none');
        setAddScheduleInputsRequired(false);
        submitBtn.textContent = `Lưu Tất Cả (${adminScheduleCart.length} môn)`;
        bulkInput.value = JSON.stringify(adminScheduleCart);

        tbody.innerHTML = adminScheduleCart.map((item, idx) => `
            <tr>
                <td class="text-center text-muted">${idx + 1}</td>
                <td><span class="badge bg-primary bg-opacity-10 text-primary border border-primary-subtle">P.${item.room}</span></td>
                <td class="text-nowrap"><strong>Thứ ${item.day_of_week}</strong> (T.${item.start_period}-${item.end_period})</td>
                <td>
                    <strong>${item.subject_name}</strong>
                    ${item.class_code ? `<span class="text-muted">(${item.class_code})</span>` : ''}
                </td>
                <td>${item.department}</td>
                <td class="text-nowrap">${item.study_period || '—'} (${item.num_weeks}T)</td>
                <td class="text-muted">${item.notes || '—'}</td>
                <td class="text-center">
                    <button type="button" class="btn btn-sm btn-outline-danger py-0 px-1" onclick="removeScheduleFromAdminCart(${idx})">&times;</button>
                </td>
            </tr>
        `).join('');
    }

    function handleAddScheduleSubmit(event) {
        const bulkInput = document.getElementById('addScheduleBulkJson');
        if (adminScheduleCart.length > 0) {
            bulkInput.value = JSON.stringify(adminScheduleCart);
            return true;
        }
        const s = parseFloat(document.getElementById('add_start_period').value);
        const e = parseFloat(document.getElementById('add_end_period').value);
        const dept = document.getElementById('add_department').value;
        if (!dept) {
            alert('Vui lòng chọn Đơn vị!');
            event.preventDefault();
            return false;
        }
        if (isNaN(s) || isNaN(e) || s >= e) {
            alert('Tiết bắt đầu phải nhỏ hơn tiết kết thúc!');
            event.preventDefault();
            return false;
        }
        return true;
    }

    function openEditScheduleModal(btnEl) {
        const id = btnEl.getAttribute('data-id');
        const room = btnEl.getAttribute('data-room') || '5.1';
        const day = btnEl.getAttribute('data-day') || '2';
        const start = btnEl.getAttribute('data-start') || '1';
        const end = btnEl.getAttribute('data-end') || '5';
        const cls = btnEl.getAttribute('data-class') || '';
        const subject = btnEl.getAttribute('data-subject') || '';
        const dept = btnEl.getAttribute('data-dept') || 'Khác';
        const period = btnEl.getAttribute('data-period') || '';
        const weeks = btnEl.getAttribute('data-weeks') || '15';
        const notes = btnEl.getAttribute('data-notes') || '';

        const form = document.getElementById('editScheduleForm');
        form.action = `/admin/edit_schedule/${id}`;

        document.getElementById('edit_room').value = room;
        document.getElementById('edit_day_of_week').value = day;
        document.getElementById('edit_start_period').value = start;
        document.getElementById('edit_end_period').value = end;
        document.getElementById('edit_class_code').value = cls;
        document.getElementById('edit_subject_name').value = subject;
        document.getElementById('edit_study_period').value = period;
        document.getElementById('edit_num_weeks').value = weeks;
        document.getElementById('edit_notes').value = notes;

        const deptSel = document.getElementById('edit_department');
        let foundOption = Array.from(deptSel.options).some(opt => opt.value === dept);
        if (!foundOption && dept) {
            const customOpt = document.createElement('option');
            customOpt.value = dept;
            customOpt.textContent = dept;
            deptSel.appendChild(customOpt);
        }
        deptSel.value = dept || 'Khác';

        const modal = new bootstrap.Modal(document.getElementById('editScheduleModal'));
        modal.show();
    }

    function handleAdminStatusDropdownChange(selectEl) {
        const actionUrl = selectEl.getAttribute('data-action') || '';
        const reqType = selectEl.getAttribute('data-type') || 'booking';
        const prevStatus = selectEl.getAttribute('data-current') || 'Processing';
        const status = selectEl.value;
        const userName = selectEl.getAttribute('data-user') || 'Giảng viên';
        const userEmail = selectEl.getAttribute('data-email') || '';
        const summary = selectEl.getAttribute('data-summary') || '';
        const existingNote = selectEl.getAttribute('data-note') || '';

        // Khôi phục lại giá trị hiển thị của select phòng trường hợp Admin đóng Modal hủy thao tác
        selectEl.value = prevStatus;

        const typeLabels = STATUS_DROPDOWN_LABELS[reqType] || {};
        const statusLabel = typeLabels[status] || 'Cập nhật trạng thái';

        currentReplyContext = { actionUrl, reqType, status, statusLabel, userName, userEmail, summary };

        const formEl = document.getElementById('statusReplyForm');
        formEl.action = actionUrl;
        document.getElementById('srHiddenStatus').value = status;
        document.getElementById('srBadgeLabel').textContent = statusLabel;
        document.getElementById('srUserLabel').textContent = userName;
        document.getElementById('srSummaryText').textContent = summary;
        document.getElementById('srEmailText').textContent = userEmail ? userEmail : 'Không có email đính kèm';
        document.getElementById('srNoteInput').value = existingNote;

        const badgeEl = document.getElementById('srBadgeLabel');
        const submitBtn = document.getElementById('srSubmitBtn');
        if (status === 'Processing') {
            badgeEl.className = 'badge bg-info text-dark rounded-pill px-3 py-1 fw-bold';
            submitBtn.className = 'btn btn-sm btn-info text-dark fw-bold px-3 shadow-sm';
            submitBtn.innerHTML = 'Xác nhận Đã tiếp nhận &amp; Gửi Email';
        } else if (status === 'Resolved') {
            badgeEl.className = 'badge bg-success rounded-pill px-3 py-1 fw-bold';
            submitBtn.className = 'btn btn-sm btn-success fw-bold px-3 shadow-sm';
            submitBtn.innerHTML = 'Xác nhận Đã xong &amp; Gửi Email';
        } else if (status === 'Approved') {
            badgeEl.className = 'badge bg-success rounded-pill px-3 py-1 fw-bold';
            submitBtn.className = 'btn btn-sm btn-success fw-bold px-3 shadow-sm';
            submitBtn.innerHTML = 'Xác nhận Đã duyệt &amp; Gửi Email';
        } else {
            badgeEl.className = 'badge bg-danger rounded-pill px-3 py-1 fw-bold';
            submitBtn.className = 'btn btn-sm btn-danger fw-bold px-3 shadow-sm';
            submitBtn.innerHTML = 'Xác nhận Từ chối &amp; Gửi Email';
        }

        const gmailBtn = document.getElementById('srGmailWebBtn');
        if (gmailBtn) {
            gmailBtn.style.display = userEmail ? 'inline-block' : 'none';
        }

        const chipsWrap = document.getElementById('srQuickChips');
        chipsWrap.innerHTML = '';
        const typeGroup = QUICK_REPLY_TEMPLATES[reqType] || {};
        const templates = typeGroup[status] || [];
        templates.forEach(tpl => {
            const b = document.createElement('button');
            b.type = 'button';
            b.className = 'btn btn-sm btn-outline-secondary rounded-pill py-0 px-2 text-start';
            b.style.fontSize = '0.71rem';
            b.textContent = '+ ' + tpl;
            b.onclick = () => {
                document.getElementById('srNoteInput').value = tpl;
            };
            chipsWrap.appendChild(b);
        });

        const modal = new bootstrap.Modal(document.getElementById('statusReplyModal'));
        modal.show();
    }

    function openGmailComposeFallback() {
        const to = (currentReplyContext.userEmail || '').trim();
        if (!to) {
            alert('Yêu cầu này không có địa chỉ email của người gửi.');
            return;
        }
        const note = (document.getElementById('srNoteInput').value || '').trim();
        let reqTypeLabel = 'đăng ký mượn phòng máy';
        if (currentReplyContext.reqType === 'issue') {
            reqTypeLabel = 'báo cáo sự cố kỹ thuật';
        } else if (currentReplyContext.reqType === 'cancellation') {
            reqTypeLabel = 'thông báo nghỉ học';
        }
        const cleanStatus = (currentReplyContext.statusLabel || '').replace(/^[^\wÀ-ỹ]+/, '').trim();
        const subject = `Phản hồi yêu cầu ${reqTypeLabel} (${currentReplyContext.summary}) - Phòng máy tính B4.2`;
        const bodyLines = [
            `Kính gửi Thầy/Cô ${currentReplyContext.userName},`,
            ``,
            `Hệ thống Quản lý Phòng máy tính Tòa B4.2 (Phòng Quản trị Thiết bị) xin thông báo kết quả xử lý yêu cầu:`,
            `- Trạng thái: ${cleanStatus}`,
            `- Chi tiết yêu cầu: ${currentReplyContext.summary}`,
            note ? `- Phản hồi từ Quản trị viên Phòng máy (QTTB): ${note}` : ``,
            ``,
            `Trân trọng,`,
            `Bộ phận quản lý phòng máy`
        ].filter(Boolean).join('\n');

        const gmailUrl = `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(to)}&su=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyLines)}`;
        window.open(gmailUrl, '_blank');
    }

    async function sendTestSmtpEmail() {
        const btn = document.getElementById('btnTestSmtp');
        const resBox = document.getElementById('ecTestResult');
        const smtpEmail = document.getElementById('ecSmtpEmail').value.trim();
        const smtpPassword = document.getElementById('ecSmtpPassword').value.trim();
        const senderName = document.getElementById('ecSenderName').value.trim();

        btn.disabled = true;
        btn.textContent = 'Đang gửi thử...';
        resBox.classList.remove('d-none', 'text-success', 'text-danger');
        resBox.textContent = 'Đang kết nối tới máy chủ smtp.gmail.com:587...';

        try {
            const resp = await fetch('/admin/email-test', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    to_email: smtpEmail,
                    smtp_email: smtpEmail,
                    smtp_password: smtpPassword,
                    smtp_sender_name: senderName
                })
            });
            const data = await resp.json();
            if (data.success) {
                resBox.classList.add('text-success');
                resBox.textContent = data.message + ' — Hãy bấm "Lưu cấu hình Email" để áp dụng!';
            } else {
                resBox.classList.add('text-danger');
                resBox.textContent = data.message;
            }
        } catch (err) {
            resBox.classList.add('text-danger');
            resBox.textContent = 'Lỗi kết nối: ' + err;
        } finally {
            btn.disabled = false;
            btn.textContent = 'Gửi Email Test';
        }
    }

    // ==========================================
    // AUTO-RELOAD SAU 30S CHO CÁC TAB XỬ LÝ GẤP
    // (Yêu cầu, Báo vắng/nghỉ, Sự cố)
    // ==========================================
    let autoReloadTimer = null;
    let autoReloadSeconds = 30;
    let isAutoReloadPaused = false;
    const URGENT_TABS = ['requests', 'cancellations', 'issues'];

    function getActiveAdminTab() {
        const activeTabEl = document.querySelector('#adminTabs .nav-link.active');
        if (activeTabEl) {
            const target = activeTabEl.getAttribute('data-bs-target') || '';
            return target.replace('#', '');
        }
        const urlParams = new URLSearchParams(window.location.search);
        return urlParams.get('tab') || 'dashboard';
    }

    function updateAutoReloadUI() {
        const badges = document.querySelectorAll('.auto-reload-badge');
        badges.forEach(b => {
            const secSpan = b.querySelector('.auto-reload-sec');
            const icon = b.querySelector('.auto-reload-icon');
            if (isAutoReloadPaused) {
                b.classList.add('paused');
                b.setAttribute('title', 'Tự động làm mới đang TẠM DỪNG. Bấm để bật lại.');
                if (secSpan) secSpan.textContent = 'Đã tạm dừng';
                if (icon) icon.className = 'bi bi-pause-circle text-warning me-1 auto-reload-icon';
            } else {
                b.classList.remove('paused');
                b.setAttribute('title', 'Tự động cập nhật sau mỗi 30s. Bấm để tạm dừng.');
                if (secSpan) secSpan.textContent = autoReloadSeconds + 's';
                if (icon) {
                    const isSpinning = autoReloadSeconds <= 2 ? ' spin-reload-active' : '';
                    icon.className = 'bi bi-arrow-repeat text-primary me-1 auto-reload-icon' + isSpinning;
                }
            }
        });
    }

    function toggleAutoReload() {
        isAutoReloadPaused = !isAutoReloadPaused;
        if (!isAutoReloadPaused) {
            autoReloadSeconds = 30;
        }
        updateAutoReloadUI();
    }

    function startAutoReloadWatcher() {
        if (autoReloadTimer) clearInterval(autoReloadTimer);
        autoReloadSeconds = 30;
        updateAutoReloadUI();

        autoReloadTimer = setInterval(() => {
            const currentTab = getActiveAdminTab();
            if (!URGENT_TABS.includes(currentTab)) {
                return; // Chỉ đếm ngược khi đang ở các tab xử lý gấp
            }

            if (isAutoReloadPaused) {
                return;
            }

            // Kiểm tra an toàn: Nếu admin đang thao tác (mở modal, nhập text, chọn checkbox xóa hàng loạt)
            const isModalOpen = document.querySelector('.modal.show');
            const activeEl = document.activeElement;
            const isInputFocused = activeEl && ['INPUT', 'TEXTAREA', 'SELECT'].includes(activeEl.tagName);
            const hasCheckedItem = document.querySelector('.item-check-requests:checked, .item-check-cancellations:checked, .item-check-issues:checked');

            // Kiểm tra xem có từ khóa tìm kiếm đang gõ hay không
            const searchInput = document.getElementById(`searchInput-${currentTab}`);
            const hasSearchText = searchInput && searchInput.value.trim().length > 0;

            if (isModalOpen || isInputFocused || hasCheckedItem || hasSearchText) {
                // Đang bận thao tác, tạm hoãn và giữ ở mức 10s để không gây gián đoạn công việc
                autoReloadSeconds = 10;
                updateAutoReloadUI();
                return;
            }

            autoReloadSeconds--;
            if (autoReloadSeconds <= 0) {
                autoReloadSeconds = 30;
                // Giữ nguyên tab hiện tại khi tải lại
                const url = new URL(window.location.href);
                url.searchParams.set('tab', currentTab);
                window.location.href = url.toString();
            } else {
                updateAutoReloadUI();
            }
        }, 1000);
    }
