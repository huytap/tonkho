let inventoryData = [];
let orderData = [];
let inventoryLoaded = false;
let orderLoaded = false;
let currentFilteredOrders = []; // Lưu kết quả tính toán hiện tại để export
let checkLocation = '';

// Lấy key đúng dù có khoảng trắng đầu/cuối
function getValueTrimmed(row, key) {
    const foundKey = Object.keys(row).find(k => k.trim() === key);
    return foundKey ? row[foundKey] : null;
}
// Xử lý upload file tồn kho
document.getElementById('inventoryFile').addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (file) {
        document.getElementById('inventoryFileName').textContent = '✓ ' + file.name;
        readInventoryFile(file);
    }
});

// Xử lý upload file đơn hàng
document.getElementById('orderFile').addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (file) {
        document.getElementById('orderFileName').textContent = '✓ ' + file.name;
        readOrderFile(file);
    }
});

// Đọc file tồn kho
function readInventoryFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellFormula: false, cellDates: true });

            // ✅ Khởi tạo mảng tạm bên trong hàm
            const allRows = [];
            workbook.SheetNames.forEach(sheetName => {
                console.log('Đang đọc sheet:', sheetName);
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
                jsonData.forEach(row => {
                    const totalWeight = getValueTrimmed(row, 'Total Weight');
                    if (row['GSM'] && totalWeight) {
                        allRows.push({ ...row, _sheet: sheetName }); // thêm tên sheet để phân biệt
                    }
                });
            });

            // ✅ Chỉ gán một lần sau khi đọc xong tất cả sheet
            inventoryData = allRows;
            inventoryLoaded = true;
            checkAndEnableCalculation();
            console.log('Đã đọc được', inventoryData.length, 'dòng tồn kho');
        } catch (error) {
            alert('Lỗi đọc file tồn kho: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Đọc file đơn hàng
function readOrderFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellFormula: false, cellDates: true });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            orderData = XLSX.utils.sheet_to_json(worksheet, { defval: null });

            orderLoaded = true;
            checkAndEnableCalculation();
            populateCustomerFilter();
            console.log('Đã đọc được', orderData.length, 'đơn hàng');
        } catch (error) {
            alert('Lỗi đọc file đơn hàng: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Kiểm tra và kích hoạt nút tính toán
function checkAndEnableCalculation() {
    if (inventoryLoaded && orderLoaded) {
        document.getElementById('filterSection').style.display = 'block';
        document.getElementById('calculateBtn').disabled = false;
    }
}

// Điền danh sách khách hàng
function populateCustomerFilter() {
    const customers = [...new Set(orderData.map(o => o['Khách hàng']).filter(c => c))].sort();
    const select = document.getElementById('customerFilter');
    select.innerHTML = '<option value="">-- Tất cả --</option>';
    customers.forEach(customer => {
        const option = document.createElement('option');
        option.value = customer;
        option.textContent = customer;
        select.appendChild(option);
    });
}

// Lấy tồn kho theo GSM
function getInventoryWeight(gsm, paperType) {
    if (!gsm) return 0;
    return inventoryData
        .filter(item => parseFloat(item['GSM']) === parseFloat(gsm) && item['Loại Giấy'] === paperType)
        .reduce((sum, item) => sum + (parseFloat(getValueTrimmed(item, 'Total Weight')) || 0), 0);
}

// Lấy chi tiết tồn kho theo GSM
function getInventoryDetails(gsm, paperType) {
    if (!gsm) return [];
    console.log(inventoryData)
    return inventoryData
        .filter(item => parseFloat(item['GSM']) === parseFloat(gsm) && getValueTrimmed(item, 'Total Weight') > 0 && item['Loại Giấy'] === paperType)
        .map(item => ({
            rollCode: item['MÃ VẬT TƯ'] || 'N/A',
            brand: item['Hiệu Giấy'] || 'N/A',
            gsm: item['GSM'],
            width: item['Kích Thước'] || 'N/A',
            weight: parseFloat(getValueTrimmed(item, 'Total Weight')) || 0,
        }));
}

function selectOptimalRoll(inventoryDetails, requiredWidth, requiredLength) {
    let allCandidates = [];
    let bestScenarioNM1 = new Map(); // Lưu phương án tốt nhất cho từng cuộn (không ghi đè cuộn khác)
    let i = 0;
    for (const roll of inventoryDetails) {
        const rollWidth = parseFloat(roll.width);
        if (!rollWidth) continue;

        // Định nghĩa các Kịch bản cắt tiềm năng
        let scenarios = [
            { width: parseFloat(requiredWidth), description: "Gốc" }
        ];

        // Nếu ở NM1, thêm Kịch bản B: Đảo chiều
        if (checkLocation === 'NM1') {
            scenarios.push({
                width: parseFloat(requiredLength),
                description: "Đảo chiều"
            });
        }

        // Tạo key duy nhất cho mỗi cuộn (VD: mã + kích thước)
        const rollKey = `${roll.rollCode || roll['MÃ VẬT TƯ'] || 'unknown'}_${rollWidth}_${roll.weight || 0}_${i++}`;

        for (const scenario of scenarios) {
            const required = scenario.width;

            // 1. Tính toán khả năng cắt
            const cuts = Math.floor(rollWidth / required);
            if (cuts === 0) continue;

            // 2. Tính Lãng phí thực tế
            const cutValue = cuts * required;
            const waste = rollWidth - cutValue;

            // 3. Xây dựng Chỉ số Quyết định (Decision Score)
            let decisionScore = waste;
            if (waste === 0) decisionScore -= 0.005;
            if (rollWidth === 60) decisionScore -= 0.01;

            // 4. Gộp kết quả
            const candidate = {
                ...roll,
                usedWidth: required,
                cutsPerRoll: cuts,
                waste,
                score: decisionScore,
                scenario: scenario.description
            };

            // 🔹 Nếu là NM1 thì chọn kịch bản tốt nhất cho từng cuộn riêng
            if (checkLocation === 'NM1') {
                if (!bestScenarioNM1.has(rollKey)) {
                    bestScenarioNM1.set(rollKey, candidate);
                } else {
                    const current = bestScenarioNM1.get(rollKey);
                    // So sánh waste → lấy phương án tốt hơn
                    if (candidate.waste < current.waste) {
                        bestScenarioNM1.set(rollKey, candidate);
                    }
                }
            } else {
                // Các location khác, thêm trực tiếp
                allCandidates.push(candidate);
            }
        }
    }

    // 🔹 Sau khi xử lý hết: lấy kịch bản tốt nhất cho NM1
    if (checkLocation === 'NM1') {
        bestScenarioNM1.forEach(candidate => {
            allCandidates.push(candidate);
        });
    }

    if (allCandidates.length === 0) return [];

    // 5. Sắp xếp theo score tốt nhất
    const sortedCandidates = [...allCandidates].sort((a, b) => a.score - b.score);
    return sortedCandidates;
}

// Tính toán trọng lượng đơn hàng
function calculateWeight(order) {
    const quantity = parseFloat(order['SL ĐH']) || 0;
    const cutWidth = parseFloat(order['Cắt tới (cm)']) || 0;
    const rollWidth = parseFloat(order['Cuồn (cm)']) || 0;
    const gsm = parseFloat(order['gsm']) || 0;
    const units = parseFloat(order['số đv']) || 1;

    if (quantity > 0 && cutWidth > 0 && rollWidth > 0 && gsm > 0) {
        return (gsm * rollWidth * cutWidth * quantity * Math.pow(10, -7)) / units;
    }
    return 0;
}

// Render kết quả
function renderResults(filteredOrders) {
    currentFilteredOrders = filteredOrders;
    const resultsDiv = document.getElementById('results');
    resultsDiv.style.display = 'block';

    // Hiển thị nút export
    document.getElementById('exportBtn').style.display = 'inline-flex';

    let html = '';

    // Thống kê
    const approvedOrders = filteredOrders.filter(o => o.type === 'approved');
    const forecastOrders = filteredOrders.filter(o => o.type === 'forecast');

    const totalWeightApproved = approvedOrders.reduce((sum, o) => sum + o.weight, 0);
    const totalWeightForecast = forecastOrders.reduce((sum, o) => sum + o.weight, 0);

    html += `
        <div class="stats">
            <div class="stat-card">
                <h3>Tổng số đơn hàng</h3>
                <div class="value">${filteredOrders.length}</div>
            </div>
            <div class="stat-card">
                <h3>Đơn hàng đã duyệt</h3>
                <div class="value">${approvedOrders.length}</div>
            </div>
            <div class="stat-card">
                <h3>Đơn hàng forecast</h3>
                <div class="value">${forecastOrders.length}</div>
            </div>
            <div class="stat-card">
                <h3>Trọng lượng tổng (kg)</h3>
                <div class="value">${(totalWeightApproved + totalWeightForecast).toLocaleString()}</div>
            </div>
        </div>
    `;

    // Render đơn hàng đã duyệt
    let globalOrderIndex = 0;
    if (approvedOrders.length > 0) {
        html += '<h2 class="section-title">📋 Đơn Hàng Đã Duyệt</h2>';
        approvedOrders.forEach(order => {
            html += renderOrderCard(order, globalOrderIndex++, 'approved');
        });
        // approvedOrders.forEach((order, idx) => {
        //     html += renderOrderCard(order, idx, 'approved');
        // });
    }

    // Render đơn hàng forecast
    if (forecastOrders.length > 0) {
        html += '<h2 class="section-title">🔮 Đơn Hàng Forecast</h2>';
        forecastOrders.forEach(order => {
            html += renderOrderCard(order, globalOrderIndex++, 'forecast');
        });
        // forecastOrders.forEach((order, idx) => {
        //     html += renderOrderCard(order, idx, 'forecast');
        // });
    }

    // Tổng kết
    if (filteredOrders.length > 0) {
        html += `
            <div class="summary-box">
                <h3>📈 Tổng Kết</h3>
                <div class="summary-grid">
                    <div class="summary-item">
                        <div class="summary-label">Tổng trọng lượng đơn đã duyệt</div>
                        <div class="summary-value" style="color: #28a745;">${totalWeightApproved.toLocaleString()} kg</div>
                    </div>
                    <div class="summary-item">
                        <div class="summary-label">Tổng trọng lượng forecast</div>
                        <div class="summary-value" style="color: #ffc107;">${totalWeightForecast.toLocaleString()} kg</div>
                    </div>
                    <div class="summary-item">
                        <div class="summary-label">Tổng cộng</div>
                        <div class="summary-value" style="color: #667eea;">${(totalWeightApproved + totalWeightForecast).toLocaleString()} kg</div>
                    </div>
                </div>
            </div>
        `;
    }

    resultsDiv.innerHTML = html;
}

// Render từng order card
function renderOrderCard(order, index, type) {
    const inventoryWeight = getInventoryWeight(order.gsm, order.paperType);
    const inventoryDetails = getInventoryDetails(order.gsm, order.paperType);
    const checkTon = selectOptimalRoll(inventoryDetails, order.rollWidth, order.cutWidth);
    const difference = inventoryWeight - order.weight;
    const typeLabel = type === 'approved' ? 'Đã Duyệt' : 'Forecast';
    const typeClass = type === 'approved' ? 'approved' : 'forecast';
    const orderLabel = type === 'approved' ? `Đơn hàng #${index + 1}: ${order.orderCode}` : `Đơn hàng Forecast #${index + 1}`;

    let html = `
        <div class="order-card">
            <div class="order-header">
                <div class="order-code">${orderLabel}</div>
                <span class="order-type ${typeClass}">${typeLabel}</span>
            </div>
            
            <div class="order-details">
                <div class="detail-item">
                    <div class="detail-label">Khách hàng</div>
                    <div class="detail-value">${order.customer}</div>
                </div>
                <div class="detail-item">
                    <div class="detail-label">Tên sản phẩm</div>
                    <div class="detail-value">${order.product}</div>
                </div>
                <div class="detail-item">
                    <div class="detail-label">GSM</div>
                    <div class="detail-value">${order.gsm}</div>
                    <div class="detail-label">Loại Giấy</div>
                    <div class="detail-value">${order.paperType}</div>
                </div>
                <div class="detail-item">
                    <div class="detail-label">Cuồn (cm)</div>
                    <div class="detail-value">${order.rollWidth}</div>
                </div>
                <div class="detail-item">
                    <div class="detail-label">Cắt tới (cm)</div>
                    <div class="detail-value">${order.cutWidth}</div>
                </div>
                <div class="detail-item">
                    <div class="detail-label">SL ĐH</div>
                    <div class="detail-value">${order.quantity.toLocaleString()}</div>
                </div>
            </div>
            
            <div class="calculation-box">
                <h4>📊 Tính toán số giấy cần sử dụng</h4>
                <div class="calculation-row">
                    <span>Trọng lượng cần (kg):</span>
                    <span>${order.weight.toLocaleString()} kg</span>
                </div>
                <div class="calculation-row">
                    <span>Tồn kho hiện tại (kg):</span>
                    <span>${inventoryWeight.toLocaleString()} kg</span>
                </div>
                <div class="calculation-row">
                    <span><strong>Chênh lệch (kg):</strong></span>
                    <span><strong style="color: ${difference >= 0 ? '#dc3545' : '#28a745'}">
                        ${difference >= 0 ? '+' : ''}${difference.toLocaleString()} kg
                    </strong></span>
                </div>
            </div>
    `;

    // Chi tiết tồn kho
    if (checkTon && checkTon.length > 0) {
        html += `
            <div class="inventory-details">
                <h4>📦 Chi tiết tồn kho & Quyết định (GSM: ${order.gsm})</h4>
                <table class="inventory-table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Mã VT</th>
                            <th>Hiệu</th>
                            <th>GSM</th>
                            <th>Rộng Cuộn (cm)</th>
                            <th>Rộng Cắt Thực Tế (cm)</th>
                            <th>Lãng Phí (cm)</th>
                            <th>Trọng lượng (kg)</th>
                            <th>Kịch Bản</th>
                            <th>Ưu Tiên</th>
                            <th>#</th>
                        </tr>
                    </thead>
                    <tbody>`;

        checkTon.forEach((detail, idx) => {
            const rowClass = idx === 0 ? 'optimal-roll' : '';
            const isOptimal = idx === 0;

            html += `
                <tr class="${rowClass}">
                    <td>${idx + 1}</td>
                    <td>${detail.rollCode}</td>
                    <td>${detail.brand}</td>
                    <td>${detail.gsm}</td>
                    <td>${detail.width}</td>
                    <td><strong class="${isOptimal ? 'highlight-value' : ''}">
                        ${detail.usedWidth ? detail.usedWidth : detail.width}
                    </strong></td>
                    <td>${detail.waste !== undefined ? detail.waste.toFixed(2) : 'N/A'}</td>
                    <td>${detail.weight.toLocaleString()}</td>
                    <td><span class="scenario-tag ${detail.scenario === 'Đảo chiều' ? 'tag-reverse' : 'tag-normal'}">
                        ${detail.scenario || 'Gốc'}
                    </span></td>
                    <td>
                        <strong class="${isOptimal ? 'optimal-label' : ''}">
                            ${isOptimal ? 'CHỌN' : 'Dự phòng'}
                        </strong>
                    </td>
                    <td>
                        <input onclick="handleCheckboxChange(this)"
                            type="checkbox"
                            class="export-checkbox"
                            data-order-index="${index}"
                            data-detail-index="${idx}"
                            ${idx === 0 ? 'checked' : ''}
                        >
                    </td>
                </tr>
            `;
        });

        html += `
                <tr class="inventory-total">
                    <td colspan="7" style="text-align: right;"><strong>Tổng Tồn Kho:</strong></td>
                    <td><strong>${inventoryWeight.toLocaleString()} kg</strong></td>
                    <td colspan="2"></td>
                </tr>
            </tbody>
        </table>
        <p class="summary-note">🎯 **Quyết định:** Chọn cuộn **${checkTon[0].rollCode}** (${checkTon[0].width}cm) với lãng phí ${checkTon[0].waste.toFixed(2)}cm (${checkTon[0].scenario} mode).</p>
        </div>`;
    } else {
        html += `<div class="alert alert-danger">❌ KHÔNG tìm thấy cuộn giấy phù hợp (${order.gsm} GSM, ${order.paperType}).</div>`;
    }

    html += '</div>';
    return html;
}

// ========== XỬ LÝ TÍNH TOÁN - QUAN TRỌNG: BỔ SUNG ĐẦY ĐỦ DỮ LIỆU CHO EXPORT ==========
document.getElementById('calculateBtn').addEventListener('click', function () {
    const customerFilter = document.getElementById('customerFilter').value;
    const gsmFilter = document.getElementById('gsmFilter').value;
    const orderTypeFilter = document.getElementById('orderTypeFilter').value;
    checkLocation = document.getElementById('locationFilter').value;

    // Filter đơn hàng
    let filtered = orderData.filter(order => {
        if (customerFilter && order['Khách hàng'] !== customerFilter) return false;
        if (gsmFilter) {
            const gsmArray = gsmFilter.split(',').map(g => g.trim());
            if (!gsmArray.includes(String(order['gsm']))) return false;
        }
        return true;
    });

    // Phân loại và tính toán - BỔ SUNG ĐẦY ĐỦ DỮ LIỆU
    const processedOrders = filtered.map(order => {
        const loaiDonHang = String(order['Loại ĐH'] || '').toLowerCase();
        const isForecast = loaiDonHang.includes('forecast') || loaiDonHang.includes('dự báo');

        if (orderTypeFilter === 'approved' && isForecast) return null;
        if (orderTypeFilter === 'forecast' && !isForecast) return null;

        const gsm = order['gsm'] || '';
        const paperType = order['Loại giấy'] || '';
        const rollWidth = order['Cuồn (cm)'] || '';
        const cutWidth = order['Cắt tới (cm)'] || '';

        // LẤY THÔNG TIN CHI TIẾT TỒN KHO VÀ CUỘN TỐI ƯU
        const inventoryDetails = getInventoryDetails(gsm, paperType);
        const checkTon = selectOptimalRoll(inventoryDetails, rollWidth, cutWidth);
        const inventoryWeight = getInventoryWeight(gsm, paperType);

        // CUỘN TỐI ƯU ĐƯỢC CHỌN (index 0)
        const selectedRoll = checkTon && checkTon.length > 0 ? checkTon[0] : null;

        return {
            type: isForecast ? 'forecast' : 'approved',
            orderCode: order['Mã DHB'] || 'N/A',
            customer: order['Khách hàng'] || 'N/A',
            product: order['Tên sản phẩm'] || 'N/A',
            gsm: gsm,
            rollWidth: rollWidth,
            cutWidth: cutWidth,
            quantity: parseFloat(order['SL ĐH']) || 0,
            units: parseFloat(order['số đv']) || 1,
            weight: calculateWeight(order),
            paperType: paperType,

            // ========== BỔ SUNG: DỮ LIỆU ĐẦY ĐỦ CHO EXPORT ==========
            inventoryWeight: inventoryWeight,
            inventoryDetails: inventoryDetails,  // Tất cả các cuộn khả dụng
            checkTon: checkTon,                  // Tất cả các cuộn đã sắp xếp theo độ ưu tiên
            selectedRoll: selectedRoll,          // Cuộn được chọn

            // Thông tin cuộn được chọn (để dễ access)
            selectedRollCode: selectedRoll ? selectedRoll.rollCode : '',
            selectedBrand: selectedRoll ? selectedRoll.brand : '',
            selectedWidth: selectedRoll ? selectedRoll.width : '',
            selectedUsedWidth: selectedRoll ? selectedRoll.usedWidth : '',
            selectedWaste: selectedRoll ? selectedRoll.waste : 0,
            selectedScenario: selectedRoll ? selectedRoll.scenario : '',
            selectedWeight: selectedRoll ? selectedRoll.weight : 0
        };
    }).filter(o => o !== null);

    renderResults(processedOrders);
});

// Xử lý export Excel
document.getElementById('exportBtn').addEventListener('click', function () {
    const exportOrders = collectExportOrders(currentFilteredOrders);
    exportToExcel(exportOrders);
});
function handleCheckboxChange(checkboxElement) {
    if (checkboxElement.checked) {
        checkboxElement.setAttribute('checked', 'checked');
    } else {
        checkboxElement.removeAttribute('checked');
    }
}
// ========== HÀM EXPORT EXCEL VỚI EXCELJS - HỖ TRỢ STYLING ĐẦY ĐỦ ==========
function collectExportOrders(orders) {
    const map = {};

    document.querySelectorAll('.export-checkbox:checked')
        .forEach(cb => {
            const orderIdx = cb.dataset.orderIndex;
            const detailIdx = cb.dataset.detailIndex;

            if (!map[orderIdx]) {
                map[orderIdx] = [];
            }
            map[orderIdx].push(parseInt(detailIdx));
        });

    // Clone dữ liệu orders theo checkbox
    return orders.map((order, index) => {
        if (!map[index]) {
            // Không tick gì → mặc định lấy dòng CHỌN
            return {
                ...order,
                checkTon: order.checkTon ? [order.checkTon[0]] : []
            };
        }

        return {
            ...order,
            checkTon: map[index].map(i => order.checkTon[i])
        };
    });
}



async function exportToExcel(orders) {
    console.log('Exporting orders:', orders);

    // Validation
    if (!orders || orders.length === 0) {
        console.log('❌ Không có dữ liệu để export!');
        return;
    }

    try {
        // Import ExcelJS từ CDN
        const ExcelJS = window.ExcelJS;
        if (!ExcelJS) {
            console.log('❌ Lỗi: Thư viện ExcelJS chưa được tải. Vui lòng kiểm tra kết nối internet.');
            return;
        }

        // Tạo workbook mới
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Tính toán tồn kho');

        // Định nghĩa các cột
        worksheet.columns = [
            { header: 'STT', key: 'stt', width: 6 },
            { header: 'KHÁCH HÀNG', key: 'khachHang', width: 12 },
            { header: 'SẢN PHẨM', key: 'tenSP', width: 16 },
            { header: 'Tổng FC (sp)', key: 'FC', width: 20 },
            //{ header: 'Tên sản phẩm', key: 'tenSP', width: 35 },
            //{ header: 'SL ĐH', key: 'slDH', width: 10 },
            { header: 'Loại giấy', key: 'loaiGiay', width: 12 },
            { header: 'DL (gsm)', key: 'gsm', width: 8 },
            { header: 'Cuồn (cm)', key: 'cuon', width: 12 },
            { header: 'Cắt tới (cm)', key: 'catToi', width: 12 },
            { header: 'Số Kg', key: 'slSuDung', width: 15 },
            { header: 'Số Tờ', key: 'slTo', width: 15 },
            { header: 'Số ĐV/tờ', key: 'soDv', width: 8 },
            { header: 'Số SP', key: 'slDH', width: 15 },
            { header: 'Hiệu', key: 'hieuG', width: 15 },
            { header: 'Tồn kho (Kg)', key: 'tonKho', width: 15 },
            { header: 'Chênh lệch (Kg)', key: 'chenhLech', width: 15 },
            { header: '---', key: 'separator', width: 10 },
            { header: 'Mã VT', key: 'maVT', width: 16 },
            { header: 'Hiệu giấy', key: 'hieuGiay', width: 20 },
            { header: 'Rộng Cuộn (cm)', key: 'rongCuon', width: 14 },
            { header: 'Rộng Cắt (cm)', key: 'rongCat', width: 14 },
            { header: 'Lãng phí (cm)', key: 'langPhi', width: 12 },
            { header: 'Kịch bản', key: 'kichBan', width: 12 },
            { header: 'Trọng lượng Cuộn (Kg)', key: 'trongLuongCuon', width: 18 }
        ];

        // Style cho HEADER (dòng tiêu đề)
        const headerRow = worksheet.getRow(1);
        headerRow.height = 30; // Tăng chiều cao để chứa text xuống dòng
        headerRow.font = { bold: true, size: 11, color: { argb: 'FF000000' } };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFB4C7E7' } // Xanh dương nhạt
        };
        headerRow.alignment = {
            vertical: 'middle',
            horizontal: 'center',
            wrapText: true  // TỰ ĐỘNG XUỐNG DÒNG
        };
        headerRow.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };

        let exportWarnings = [];
        let currentRow = 2; // Bắt đầu từ dòng 2 (sau header)

        // Duyệt qua từng đơn hàng
        orders.forEach((order, index) => {
            try {
                const difference = order.inventoryWeight - order.weight;

                // Dòng thông tin đơn hàng chính
                const mainRow = worksheet.addRow({
                    stt: index + 1,
                    //loaiDH: order.type === 'approved' ? 'Đã duyệt' : 'Forecast',
                    khachHang: order.customer || 'N/A',
                    tenSP: order.product || 'N/A',
                    FC: '',
                    //maDHB: order.orderCode || 'N/A',
                    //slDH: order.quantity || 0,
                    loaiGiay: order.paperType || 'N/A',
                    gsm: order.gsm || 'N/A',
                    cuon: order.rollWidth || 'N/A',
                    catToi: order.cutWidth || 'N/A',
                    slSuDung: order.weight ? parseFloat(order.weight.toFixed(2)) : 0,
                    slTo: order.weight ? (order.weight / (order.gsm * order.rollWidth * order.cutWidth * 0.0000001)) : 0,
                    soDv: order.units || 1,
                    slDH: order.quantity || 0,
                    hieuG: order.selectedBrand || 'N/A',
                    tonKho: parseFloat(order.inventoryWeight.toFixed(2)),
                    chenhLech: parseFloat(difference.toFixed(2)),
                    separator: '===',
                    maVT: '',
                    hieuGiay: '',
                    rongCuon: '',
                    rongCat: '',
                    langPhi: '',
                    kichBan: '',
                    trongLuongCuon: ''
                });

                // Style cho dòng chính
                mainRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                mainRow.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                };
                // Thêm các dòng chi tiết tồn kho
                if (order.checkTon && order.checkTon.length > 0) {
                    order.checkTon.forEach((detail, idx) => {
                        const detailRow = worksheet.addRow({
                            stt: '',
                            khachHang: '',
                            tenSP: '',
                            FC: '',
                            loaiGiay: '',
                            gsm: '',
                            cuon: '',
                            catToi: '',
                            slSuDung: '',
                            slTo: '',
                            soDv: '',
                            slDH: '',
                            hieuG: '',
                            tonKho: '',
                            chenhLech: '',
                            separator: idx === 0 ? '→ CHỌN' : '→ Dự phòng',
                            maVT: detail.rollCode || 'N/A',
                            hieuGiay: detail.brand || 'N/A',
                            rongCuon: detail.width || 'N/A',
                            rongCat: detail.usedWidth || detail.width || 'N/A',
                            langPhi: detail.waste !== undefined ? parseFloat(detail.waste.toFixed(2)) : 0,
                            kichBan: detail.scenario || 'Gốc',
                            trongLuongCuon: detail.weight ? parseFloat(detail.weight.toFixed(2)) : 0
                        });

                        // Style cho dòng chi tiết
                        detailRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                        detailRow.border = {
                            top: { style: 'thin', color: { argb: 'FF000000' } },
                            left: { style: 'thin', color: { argb: 'FF000000' } },
                            bottom: { style: 'thin', color: { argb: 'FF000000' } },
                            right: { style: 'thin', color: { argb: 'FF000000' } }
                        };

                        // Highlight cho dòng được chọn
                        if (idx === 0) {
                            detailRow.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: 'FFFFEB9C' } // Vàng nhạt
                            };
                            detailRow.font = { bold: true };
                        }
                    });
                }/* else {
                    exportWarnings.push(`Đơn #${index + 1} (${order.orderCode}): Không tìm thấy cuộn phù hợp`);

                    const noRollRow = worksheet.addRow({
                        stt: '', khachHang: '', tenSP: '', FC: '',
                        loaiGiay: '', gsm: '', cuon: '', catToi: '',
                        slSuDung: '', slTo: '', soDv: '', slDH: '', tonKho: '', chenhLech: '',
                        separator: '→', maVT: '❌ KHÔNG TÌM THẤY', hieuGiay: '',
                        rongCuon: '', rongCat: '', langPhi: '', kichBan: '', trongLuongCuon: ''
                    });

                    noRollRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                    noRollRow.border = {
                        top: { style: 'thin', color: { argb: 'FF000000' } },
                        left: { style: 'thin', color: { argb: 'FF000000' } },
                        bottom: { style: 'thin', color: { argb: 'FF000000' } },
                        right: { style: 'thin', color: { argb: 'FF000000' } }
                    };
                    noRollRow.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFC7CE' } // Đỏ nhạt
                    };
                }*/

                // Dòng trống phân cách
                const emptyRow = worksheet.addRow({});
                emptyRow.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                };

            } catch (error) {
                console.error(`Lỗi khi xử lý đơn hàng #${index + 1}:`, error);
                exportWarnings.push(`Đơn #${index + 1}: Lỗi xử lý - ${error.message}`);
            }
        });

        // Tạo tên file với timestamp
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
        const filename = `TinhToanTonKho_${timestamp}.xlsx`;

        // Xuất file
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = filename;
        anchor.click();
        window.URL.revokeObjectURL(url);

        // Hiển thị thông báo thành công
        let successMsg = `✅ Export thành công!\nĐã xuất ${orders.length} đơn hàng ra file: ${filename}`;

        if (exportWarnings.length > 0) {
            successMsg += `\n\n⚠️ Lưu ý:\n${exportWarnings.join('\n')}`;
        }

        console.log(successMsg);

    } catch (error) {
        console.error('Lỗi khi tạo file Excel:', error);
        console.log(`❌ Lỗi khi tạo file Excel: ${error.message}\n\nVui lòng thử lại hoặc kiểm tra console để biết chi tiết.`);
    }
}