<!DOCTYPE html>
<html lang="vi">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ Thống Tính Toán Tồn Kho Giấy</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }

        .header h1 {
            font-size: 2em;
            margin-bottom: 10px;
        }

        .filter-section {
            padding: 30px;
            background: #f8f9fa;
            border-bottom: 3px solid #667eea;
        }

        .filter-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }

        .filter-group {
            display: flex;
            flex-direction: column;
        }

        .filter-group label {
            font-weight: bold;
            margin-bottom: 5px;
            color: #333;
        }

        .filter-group input,
        .filter-group select {
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 5px;
            font-size: 1em;
        }

        .filter-group input:focus,
        .filter-group select:focus {
            outline: none;
            border-color: #667eea;
        }

        .calc-button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 15px 30px;
            font-size: 1.2em;
            border-radius: 5px;
            cursor: pointer;
            transition: transform 0.2s;
            width: 100%;
        }

        .calc-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }

        .content {
            padding: 30px;
        }

        .stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .stat-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        }

        .stat-card h3 {
            font-size: 0.9em;
            opacity: 0.9;
            margin-bottom: 10px;
            text-transform: uppercase;
        }

        .stat-card .value {
            font-size: 2em;
            font-weight: bold;
        }

        .section-title {
            font-size: 1.5em;
            margin: 30px 0 20px 0;
            color: #333;
            border-bottom: 3px solid #667eea;
            padding-bottom: 10px;
        }

        .order-card {
            background: #f8f9fa;
            border: 1px solid #dee2e6;
            border-left: 4px solid #667eea;
            padding: 20px;
            margin-bottom: 20px;
            border-radius: 8px;
            transition: transform 0.2s, box-shadow 0.2s;
        }

        .order-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.1);
        }

        .order-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }

        .order-code {
            font-size: 1.3em;
            font-weight: bold;
            color: #667eea;
        }

        .order-type {
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: bold;
        }

        .approved {
            background: #28a745;
            color: white;
        }

        .forecast {
            background: #ffc107;
            color: #333;
        }

        .order-details {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 15px;
        }

        .detail-item {
            background: white;
            padding: 10px;
            border-radius: 5px;
        }

        .detail-label {
            font-size: 0.85em;
            color: #666;
            margin-bottom: 5px;
        }

        .detail-value {
            font-size: 1.1em;
            font-weight: bold;
            color: #333;
        }

        .calculation-box {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            margin-top: 15px;
        }

        .calculation-box h4 {
            margin-bottom: 15px;
            font-size: 1.1em;
        }

        .calculation-row {
            display: flex;
            justify-content: space-between;
            padding: 10px 0;
            border-bottom: 1px solid rgba(255, 255, 255, 0.3);
        }

        .calculation-row:last-child {
            border-bottom: none;
            font-size: 1.2em;
            font-weight: bold;
            margin-top: 10px;
            padding-top: 15px;
        }

        .no-orders {
            text-align: center;
            padding: 50px;
            color: #666;
        }

        .footer {
            text-align: center;
            padding: 20px;
            color: #666;
            font-size: 0.9em;
        }

        @media print {
            .filter-section {
                display: none;
            }

            .calc-button {
                display: none;
            }
        }
    </style>
</head>

<body>
    <?php
    require 'vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\IOFactory;

    // Lấy filter từ form
    $customerFilter = $_GET['customer'] ?? '';
    $gsmFilter = $_GET['gsm'] ?? '';
    $orderTypeFilter = $_GET['order_type'] ?? 'all';
    $location = $_GET['location'] ?? 'NM1';
    $calculate = isset($_GET['calculate']);

    class TonKhoProcessor
    {
        private $inventoryFile;
        private $orderFile;
        private $inventory = [];
        private $orders = [];

        public function __construct($inventoryFile, $orderFile)
        {
            $this->inventoryFile = $inventoryFile;
            $this->orderFile = $orderFile;
        }

        public function readInventory()
        {
            try {
                // Bật tính toán công thức Excel
                $reader = IOFactory::createReader('Xlsx');
                $reader->setReadDataOnly(false);
                $spreadsheet = $reader->load($this->inventoryFile);

                $sheets = $spreadsheet->getAllSheets();

                foreach ($sheets as $worksheet) {
                    $highestRow = $worksheet->getHighestRow();
                    $highestColumn = $worksheet->getHighestColumn();

                    $headers = [];
                    for ($col = 'A'; $col <= $highestColumn; $col++) {
                        $value = $worksheet->getCell($col . '1')->getValue();
                        $headers[] = $value;
                    }

                    for ($row = 2; $row <= $highestRow; $row++) {
                        $rowData = [];
                        $idx = 0;
                        for ($col = 'A'; $col <= $highestColumn; $col++) {
                            // Đọc cả giá trị đã tính toán
                            $value = $worksheet->getCell($col . $row)->getCalculatedValue();
                            $rowData[$headers[$idx] ?? $col] = $value;
                            $idx++;
                        }

                        // Chỉ lưu dòng có dữ liệu GSM hoặc Total Weight
                        if (!empty($rowData['GSM']) || !empty($rowData['Total Weight'])) {
                            $this->inventory[] = $rowData;
                        }
                    }
                }
                return $this->inventory;
            } catch (\Exception $e) {
                return [];
            }
        }

        // Lấy tồn kho dựa trên GSM
        public function getInventoryWeight($gsm)
        {
            if (empty($gsm)) return 0;

            $totalWeight = 0;

            foreach ($this->inventory as $item) {
                $itemGsm = $item['GSM'] ?? '';
                $itemWeight = $item['Total Weight'] ?? 0;

                // So sánh GSM và cộng tổng trọng lượng
                if (!empty($itemGsm) && !empty($itemWeight)) {
                    if (floatval($itemGsm) == floatval($gsm)) {
                        $weight = floatval($itemWeight);
                        if ($weight > 0) {
                            $totalWeight += $weight;
                        }
                    }
                }
            }

            return $totalWeight;
        }
        // Lấy danh sách chi tiết các cuộn giấy theo GSM
        public function getInventoryDetails($gsm)
        {
            if (empty($gsm)) return [];

            $details = [];

            foreach ($this->inventory as $item) {
                $itemGsm = $item['GSM'] ?? '';
                $itemWeight = $item['Total Weight'] ?? 0;

                // So sánh GSM
                if (!empty($itemGsm) && !empty($itemWeight)) {
                    if (floatval($itemGsm) == floatval($gsm)) {
                        $weight = floatval($itemWeight);
                        print_r($weight);
                        if ($weight > 0) {
                            $details[] = [
                                'brand' => $item['Hiệu giấy'] ?? 'N/A',
                                'gsm' => $itemGsm,
                                'width' => $item['Kích thước'] ?? 'N/A',
                                'roll_code' => $item['MÃ VẬT TƯ'] ?? 'N/A',
                                'weight' => $weight
                            ];
                        }
                    }
                }
            }

            return $details;
        }
        // Tính tổng tồn kho cho các đơn hàng
        public function calculateTotalInventory($orders)
        {
            $totalInventory = 0;
            // TODO: Implement mapping với tồn kho thực tế
            return $totalInventory;
        }

        public function readOrders()
        {
            try {
                $spreadsheet = IOFactory::load($this->orderFile);
                $worksheet = $spreadsheet->getActiveSheet();

                $highestRow = $worksheet->getHighestRow();
                $highestColumn = $worksheet->getHighestColumn();

                $headers = [];
                for ($col = 'A'; $col <= $highestColumn; $col++) {
                    $value = $worksheet->getCell($col . '1')->getValue();
                    $headers[] = $value;
                }

                for ($row = 2; $row <= $highestRow; $row++) {
                    $rowData = [];
                    $idx = 0;
                    for ($col = 'A'; $col <= $highestColumn; $col++) {
                        $value = $worksheet->getCell($col . $row)->getValue();
                        $rowData[$headers[$idx] ?? $col] = $value;
                        $idx++;
                    }

                    if (!empty(array_filter($rowData, function ($v) {
                        return !is_null($v) && $v !== '';
                    }))) {
                        $this->orders[] = $rowData;
                    }
                }

                return $this->orders;
            } catch (\Exception $e) {
                return [];
            }
        }

        public function getUniqueCustomers()
        {
            $customers = [];
            foreach ($this->orders as $order) {
                $customer = $order['Khách hàng'] ?? '';
                if (!empty($customer) && !in_array($customer, $customers)) {
                    $customers[] = $customer;
                }
            }
            sort($customers);
            return $customers;
        }

        public function filterAndProcess($customerFilter, $gsmFilter, $orderTypeFilter, $location)
        {
            $filteredOrders = [];

            foreach ($this->orders as $order) {
                // Filter khách hàng
                if (!empty($customerFilter) && ($order['Khách hàng'] ?? '') !== $customerFilter) {
                    continue;
                }

                // Filter GSM
                $orderGsm = $order['gsm'] ?? '';
                if (!empty($gsmFilter) && !empty($orderGsm)) {
                    $gsmArray = array_map('trim', explode(',', $gsmFilter));
                    if (!in_array($orderGsm, $gsmArray)) {
                        continue;
                    }
                }

                $filteredOrders[] = $order;
            }

            // Phân loại đơn hàng
            $approvedOrders = [];
            $forecastOrders = [];

            foreach ($filteredOrders as $order) {
                $loaiDonHang = strtolower(trim($order['Loại ĐH'] ?? ''));

                // Kiểm tra theo cột "Loại ĐH"
                if (stripos($loaiDonHang, 'forecast') !== false || stripos($loaiDonHang, 'dự báo') !== false) {
                    // Forecast order
                    if ($orderTypeFilter === 'all' || $orderTypeFilter === 'forecast') {
                        $forecastOrders[] = $order;
                    }
                } else {
                    // Approved order
                    if ($orderTypeFilter === 'all' || $orderTypeFilter === 'approved') {
                        $approvedOrders[] = $order;
                    }
                }
            }

            return [
                'approved' => $approvedOrders,
                'forecast' => $forecastOrders
            ];
        }
    }

    // Khởi tạo processor
    $processor = new TonKhoProcessor(
        'data/TỒN GIẤY CUỘN 16-10-2025.xlsx',
        'data/DON HANG.xlsx'
    );

    $inventory = $processor->readInventory();
    $orders = $processor->readOrders();
    // Lấy danh sách khách hàng
    $customers = $processor->getUniqueCustomers();

    // Nếu click nút tính toán thì hiển thị kết quả
    if ($calculate) {
        $categorized = $processor->filterAndProcess($customerFilter, $gsmFilter, $orderTypeFilter, $location);
        $approvedOrders = $categorized['approved'];
        $forecastOrders = $categorized['forecast'];

        // Tính tổng trọng lượng
        $totalWeightApproved = 0;
        $totalWeightForecast = 0;

        foreach ($approvedOrders as $order) {
            $quantity = $order['SL ĐH'] ?? 0;
            $cutWidth = $order['Cắt tới (cm)'] ?? 0;
            $rollWidth = $order['Cuồn (cm)'] ?? 0;
            $gsm = $order['gsm'] ?? 0;

            if ($quantity > 0 && $cutWidth > 0 && $rollWidth > 0 && $gsm > 0) {
                $units = floatval($order['số đv'] ?? 1);
                $weight = (floatval($gsm) * floatval($rollWidth) * floatval($cutWidth) * floatval($quantity) * pow(10, -7)) / $units;
                $totalWeightApproved += $weight;
            }
        }

        foreach ($forecastOrders as $order) {
            $quantity = $order['SL ĐH'] ?? 0;
            $cutWidth = $order['Cắt tới (cm)'] ?? 0;
            $rollWidth = $order['Cuồn (cm)'] ?? 0;
            $gsm = $order['gsm'] ?? 0;

            if ($quantity > 0 && $cutWidth > 0 && $rollWidth > 0 && $gsm > 0) {
                $units = floatval($order['số đv'] ?? 1);
                $weight = (floatval($gsm) * floatval($rollWidth) * floatval($cutWidth) * floatval($quantity) * pow(10, -7)) / $units;
                $totalWeightForecast += $weight;
            }
        }
    }
    ?>

    <div class="container">
        <div class="header">
            <h1>🏭 Hệ Thống Tính Toán Tồn Kho Giấy</h1>
            <p>Quản lý và tính toán số lượng giấy cần dùng cho đơn hàng</p>
        </div>

        <div class="filter-section">
            <form method="GET" action="">
                <div class="filter-grid">
                    <div class="filter-group">
                        <label>Khách hàng:</label>
                        <select name="customer">
                            <option value="">-- Tất cả --</option>
                            <?php foreach ($customers as $customer): ?>
                                <option value="<?php echo htmlspecialchars($customer); ?>" <?php echo $customerFilter === $customer ? 'selected' : ''; ?>>
                                    <?php echo htmlspecialchars($customer); ?>
                                </option>
                            <?php endforeach; ?>
                        </select>
                    </div>

                    <div class="filter-group">
                        <label>Định lượng (GSM):</label>
                        <input type="text" name="gsm" value="<?php echo htmlspecialchars($gsmFilter); ?>" placeholder="VD: 160, 180, 200">
                    </div>

                    <div class="filter-group">
                        <label>Loại đơn hàng:</label>
                        <select name="order_type">
                            <option value="all" <?php echo $orderTypeFilter === 'all' ? 'selected' : ''; ?>>Tất cả</option>
                            <option value="approved" <?php echo $orderTypeFilter === 'approved' ? 'selected' : ''; ?>>Đã duyệt</option>
                            <option value="forecast" <?php echo $orderTypeFilter === 'forecast' ? 'selected' : ''; ?>>Forecast</option>
                        </select>
                    </div>

                    <div class="filter-group">
                        <label>Location:</label>
                        <select name="location">
                            <option value="NM1" <?php echo $location === 'NM1' ? 'selected' : ''; ?>>NM1</option>
                            <option value="NM2" <?php echo $location === 'NM2' ? 'selected' : ''; ?>>NM2</option>
                        </select>
                    </div>
                </div>

                <button type="submit" name="calculate" value="1" class="calc-button">
                    📊 Tính Toán
                </button>
            </form>
        </div>

        <?php if ($calculate): ?>
            <div class="content">
                <!-- Thống kê -->
                <div class="stats">
                    <div class="stat-card">
                        <h3>Tổng số đơn hàng</h3>
                        <div class="value"><?php echo count($approvedOrders) + count($forecastOrders); ?></div>
                    </div>
                    <div class="stat-card">
                        <h3>Đơn hàng đã duyệt</h3>
                        <div class="value"><?php echo count($approvedOrders); ?></div>
                    </div>
                    <div class="stat-card">
                        <h3>Đơn hàng forecast</h3>
                        <div class="value"><?php echo count($forecastOrders); ?></div>
                    </div>
                    <div class="stat-card">
                        <h3>Trọng lượng tổng (kg)</h3>
                        <div class="value"><?php echo number_format($totalWeightApproved + $totalWeightForecast, 2); ?></div>
                    </div>
                </div>

                <!-- Đơn hàng đã duyệt -->
                <?php if (!empty($approvedOrders)): ?>
                    <h2 class="section-title">📋 Đơn Hàng Đã Duyệt</h2>
                    <?php foreach ($approvedOrders as $index => $order):
                        $orderCode = $order['Mã DHB'] ?? 'N/A';
                        $gsm = $order['gsm'] ?? '';
                        $rollWidth = $order['Cuồn (cm)'] ?? '';
                        $cutWidth = $order['Cắt tới (cm)'] ?? '';
                        $quantity = $order['SL ĐH'] ?? '';
                        $units = $order['số đv'] ?? '1';
                        $customer = $order['Khách hàng'] ?? 'N/A';
                        $product = $order['Tên sản phẩm'] ?? 'N/A';
                        $totalWeightInventory = $order['Total Weight'] ?? 0;
                        // Tính toán
                        $totalWeight = 0;
                        if (!empty($quantity) && !empty($cutWidth) && !empty($rollWidth) && !empty($gsm)) {
                            $totalWeight = (floatval($gsm) * floatval($rollWidth) * floatval($cutWidth) * floatval($quantity) * pow(10, -7)) / floatval($units);
                        }
                    ?>
                        <div class="order-card">
                            <div class="order-header">
                                <div class="order-code">Đơn hàng #<?php echo $index + 1; ?>: <?php echo htmlspecialchars($orderCode); ?></div>
                                <span class="order-type approved">Đã Duyệt</span>
                            </div>

                            <div class="order-details">
                                <div class="detail-item">
                                    <div class="detail-label">Khách hàng</div>
                                    <div class="detail-value"><?php echo htmlspecialchars($customer); ?></div>
                                </div>
                                <div class="detail-item">
                                    <div class="detail-label">Tên sản phẩm</div>
                                    <div class="detail-value"><?php echo htmlspecialchars($product); ?></div>
                                </div>
                                <div class="detail-item">
                                    <div class="detail-label">GSM</div>
                                    <div class="detail-value"><?php echo $gsm; ?></div>
                                </div>
                                <div class="detail-item">
                                    <div class="detail-label">Cuộn (cm)</div>
                                    <div class="detail-value"><?php echo $rollWidth; ?></div>
                                </div>
                                <div class="detail-item">
                                    <div class="detail-label">Cắt tới (cm)</div>
                                    <div class="detail-value"><?php echo $cutWidth; ?></div>
                                </div>
                                <div class="detail-item">
                                    <div class="detail-label">SL ĐH</div>
                                    <div class="detail-value"><?php echo number_format($quantity); ?></div>
                                </div>
                                <div class="detail-item">
                                    <div class="detail-label">Số đơn vị</div>
                                    <div class="detail-value"><?php echo $units; ?></div>
                                </div>
                            </div>

                            <?php
                            // Lấy tồn kho (placeholder)
                            $inventoryWeight = $processor->getInventoryWeight($gsm);
                            $difference = $totalWeight - $inventoryWeight;
                            ?>
                            <div class="calculation-box">
                                <h4>📊 Tính toán số giấy cần sử dụng</h4>
                                <div class="calculation-row">
                                    <span>Trọng lượng cần (kg):</span>
                                    <span><?php echo number_format($totalWeight, 2); ?> kg</span>
                                </div>
                                <div class="calculation-row">
                                    <span>Tồn kho hiện tại (kg):</span>
                                    <span><?php echo number_format($totalWeightInventory, 2); ?> kg</span>
                                    <?php $inventoryDetails = $processor->getInventoryDetails($gsm); ?>
                                    <pre>
                                        <?php print_r($inventoryDetails); ?>
                                    </pre>
                                </div>
                                <div class="calculation-row">
                                    <span><strong>Chênh lệch (kg):</strong></span>
                                    <span><strong style="color: <?php echo $difference >= 0 ? '#dc3545' : '#28a745'; ?>">
                                            <?php echo $difference >= 0 ? '+' : ''; ?><?php echo number_format($difference, 2); ?> kg
                                        </strong></span>
                                </div>
                            </div>
                        </div>
                    <?php endforeach; ?>
                <?php endif; ?>

                <!-- Đơn hàng forecast -->
                <?php if (!empty($forecastOrders)): ?>
                    <h2 class="section-title">🔮 Đơn Hàng Forecast</h2>
                    <?php foreach ($forecastOrders as $index => $order):
                        $gsm = $order['gsm'] ?? '';
                        $rollWidth = $order['Cuồn (cm)'] ?? '';
                        $cutWidth = $order['Cắt tới (cm)'] ?? '';
                        $quantity = $order['SL ĐH'] ?? '';

                        // Tính toán
                        $totalWeight = 0;
                        if (!empty($quantity) && !empty($cutWidth) && !empty($rollWidth) && !empty($gsm)) {
                            $units = floatval($order['số đv'] ?? 1);
                            $totalWeight = (floatval($gsm) * floatval($rollWidth) * floatval($cutWidth) * floatval($quantity) * pow(10, -7)) / $units;
                        }

                        // Lấy tồn kho (placeholder)
                        $inventoryWeight = $processor->getInventoryWeight($gsm);
                        $difference = $totalWeight - $inventoryWeight;
                    ?>
                        <div class="order-card">
                            <div class="order-header">
                                <div class="order-code">Đơn hàng Forecast #<?php echo $index + 1; ?></div>
                                <span class="order-type forecast">Forecast</span>
                            </div>

                            <div class="order-details">
                                <div class="detail-item">
                                    <div class="detail-label">Tên sản phẩm</div>
                                    <div class="detail-value"><?php echo htmlspecialchars($order['Tên sản phẩm'] ?? 'N/A'); ?></div>
                                </div>
                                <div class="detail-item">
                                    <div class="detail-label">GSM</div>
                                    <div class="detail-value"><?php echo $gsm; ?></div>
                                </div>
                                <div class="detail-item">
                                    <div class="detail-label">SL ĐH</div>
                                    <div class="detail-value"><?php echo number_format($quantity); ?></div>
                                </div>
                                <div class="detail-item">
                                    <div class="detail-label">Cắt tới (cm)</div>
                                    <div class="detail-value"><?php echo $cutWidth; ?></div>
                                </div>
                            </div>

                            <div class="calculation-box">
                                <h4>📊 Tính toán số giấy cần sử dụng (Dự báo)</h4>
                                <div class="calculation-row">
                                    <span>Trọng lượng cần (kg):</span>
                                    <span><?php echo number_format($totalWeight, 2); ?> kg</span>
                                </div>
                                <div class="calculation-row">
                                    <span>Tồn kho hiện tại (kg):</span>
                                    <span><?php echo number_format($inventoryWeight, 2); ?> kg</span>

                                </div>
                                <div class="calculation-row">
                                    <span><strong>Chênh lệch (kg):</strong></span>
                                    <span><strong style="color: <?php echo $difference >= 0 ? '#dc3545' : '#28a745'; ?>">
                                            <?php echo $difference >= 0 ? '+' : ''; ?><?php echo number_format($difference, 2); ?> kg
                                        </strong></span>
                                </div>
                            </div>
                        </div>
                    <?php endforeach; ?>
                <?php endif; ?>

                <?php if (!empty($approvedOrders) || !empty($forecastOrders)): ?>
                    <div style="margin-top: 30px; background: #f8f9fa; padding: 20px; border-radius: 8px;">
                        <h3 style="color: #333; margin-bottom: 15px;">📈 Tổng Kết</h3>
                        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
                            <div style="background: white; padding: 15px; border-radius: 5px;">
                                <div style="color: #666; font-size: 0.9em;">Tổng trọng lượng đơn đã duyệt</div>
                                <div style="font-size: 1.5em; font-weight: bold; color: #28a745; margin-top: 5px;"><?php echo number_format($totalWeightApproved, 2); ?> kg</div>
                            </div>
                            <div style="background: white; padding: 15px; border-radius: 5px;">
                                <div style="color: #666; font-size: 0.9em;">Tổng trọng lượng forecast</div>
                                <div style="font-size: 1.5em; font-weight: bold; color: #ffc107; margin-top: 5px;"><?php echo number_format($totalWeightForecast, 2); ?> kg</div>
                            </div>
                            <div style="background: white; padding: 15px; border-radius: 5px;">
                                <div style="color: #666; font-size: 0.9em;">Tổng cộng</div>
                                <div style="font-size: 1.5em; font-weight: bold; color: #667eea; margin-top: 5px;"><?php echo number_format($totalWeightApproved + $totalWeightForecast, 2); ?> kg</div>
                            </div>
                        </div>
                    </div>
                <?php endif; ?>
            </div>
        <?php endif; ?>

        <div class="footer">
            <p>© 2025 Hệ Thống Tính Toán Tồn Kho Giấy | <?php echo $location; ?> | Location-based calculation</p>
        </div>
    </div>
</body>

</html>