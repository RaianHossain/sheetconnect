<?php
/**
 * Plugin Name: SheetConnect
 * Description: This plugin connects WooCommerce orders with a Google Spreadsheet. Every order data is populated to the sheet, allowing multiple people to work on it collaboratively.
 * Version: 1.0.0
 * Author: Md. Raian Hossain
 */


if (!defined('ABSPATH')) {
    exit;
}

require_once __DIR__ . '/vendor/autoload.php'; 

require_once __DIR__ . '/includes/settings.php';

add_action('woocommerce_thankyou', 'sheetconnect_sync_order_to_google_sheet', 10, 1);
add_action('woocommerce_order_status_changed', 'sheetconnect_update_order_status_in_sheet', 10, 4);
add_action('wp_ajax_sheetconnect_sync_orders', 'sheetconnect_sync_orders_ajax_handler');

class SheetConnect_GoogleSheet {
    private $service;
    private $spreadsheetId;

    public function __construct() {
        // Get Spreadsheet ID and JSON file path from settings
        $this->spreadsheetId = get_option('sheetconnect_google_spreadsheet_id');
        $json_key_path = dirname(__DIR__, 2) . '/uploads' . '/sheetconnect' . '/' . get_option('sheetconnect_service_account_file_name');

        error_log('Initializing Google Sheets API with Spreadsheet ID: ' . $this->spreadsheetId);

        if ($this->spreadsheetId && file_exists($json_key_path)) {
            $this->authenticate($json_key_path);
        } else {
            error_log('Spreadsheet ID or JSON key file not found.');
        }
    }

    private function authenticate($json_key_path) {
        $client = new Google_Client();
        $client->setAuthConfig($json_key_path);
        $client->addScope(Google_Service_Sheets::SPREADSHEETS);
        $this->service = new Google_Service_Sheets($client);
    }

    public function write_order_to_sheet($order_id) {
        try {
            $order = wc_get_order($order_id);
            if (!$order) {
                return;
            }

            // Check if the sheet is empty; if so, initialize headers
            $sheetRange = 'Sheet1!A1:K1';
            $response = $this->service->spreadsheets_values->get($this->spreadsheetId, $sheetRange);
            if (empty($response->getValues())) {
                $headers = [['Order ID', 'Order Date', 'Customer Name', 'Address', 'Phone', 'Status', 'Order Quantity', 'Price per Piece', 'Total Amount',  'Product ID', 'Product Name']];
                $this->service->spreadsheets_values->update(
                    $this->spreadsheetId,
                    $sheetRange,
                    new Google_Service_Sheets_ValueRange(['values' => $headers]),
                    ['valueInputOption' => 'RAW']
                );

                $this->format_headers();
            }

            // Prepare batch update requests for appending order data and setting dropdowns
            $requests = [];
            $products = $order->get_items();
            foreach ($products as $product) {
                $row = [
                    $order->get_id(),
                    $order->get_date_created()->date('Y-m-d H:i:s'),
                    $order->get_billing_first_name() . ' ' . $order->get_billing_last_name(),
                    $order->get_billing_address_1(),
                    $order->get_billing_phone(),
                    $order->get_status(),
                    $product->get_quantity(),
                    number_format($order->get_total() / $product->get_quantity()),
                    number_format($order->get_total(), 2),
                    $product->get_product_id(),
                    $product->get_name(),
                ];

                $startRowIndex = $this->get_next_empty_row_index('Sheet1') - 1; // Find next empty row index

                // Append order data to the sheet
                $requests[] = new Google_Service_Sheets_Request([
                    'appendCells' => [
                        'sheetId' => 0, // Adjust sheetId if needed
                        'rows' => [['values' => $this->create_row_data($row)]],
                        'fields' => 'userEnteredValue'
                    ]
                ]);

                // Set dropdown for Order Status
                $order_status_options = ['pending', 'processing', 'on-hold', 'completed', 'cancelled', 'refunded', 'failed'];
                $requests[] = new Google_Service_Sheets_Request([
                    'setDataValidation' => [
                        'range' => [
                            'sheetId' => 0,
                            'startRowIndex' => $startRowIndex,
                            'endRowIndex' => $startRowIndex + 1,
                            'startColumnIndex' => 5, // Column for "Order Status"
                            'endColumnIndex' => 6
                        ],
                        'rule' => [
                            'condition' => [
                                'type' => 'ONE_OF_LIST',
                                'values' => array_map(function ($option) {
                                    return ['userEnteredValue' => $option];
                                }, $order_status_options),
                            ],
                            'showCustomUi' => true,
                            'strict' => true,
                        ],
                    ],
                ]);
            }

            // Batch Update API call to execute all requests in one call
            $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
                'requests' => $requests,
            ]);
            $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $batchUpdateRequest);

        } catch (Exception $e) {
            error_log('Error writing order to sheet: ' . $e->getMessage());
        }
    }

    private function create_row_data($row) {
        $rowData = [];
        foreach ($row as $cellValue) {
            $rowData[] = new Google_Service_Sheets_CellData([
                'userEnteredValue' => ['stringValue' => (string)$cellValue]
            ]);
        }
        return $rowData;
    }

    private function get_next_empty_row_index($sheetName) {
        $range = $sheetName . '!A:A'; // Assuming column A is always filled with Order ID or some data
        $response = $this->service->spreadsheets_values->get($this->spreadsheetId, $range);
        $values = $response->getValues();
        return count($values) + 1; // Return the next empty row index
    }

    private function format_headers() {
        $requests = [
            'repeatCell' => [
                'range' => ['sheetId' => 0, 'startRowIndex' => 0, 'endRowIndex' => 1],
                'cell' => [
                    'userEnteredFormat' => [
                        'backgroundColor' => ['red' => 0.2, 'green' => 0.7, 'blue' => 0.2],
                        'textFormat' => ['bold' => true, 'fontSize' => 12]
                    ]
                ],
                'fields' => 'userEnteredFormat(backgroundColor,textFormat)'
            ]
        ];

        $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => [$requests]
        ]);

        $this->service->spreadsheets->batchUpdate($this->spreadsheetId, $batchUpdateRequest);
    }

    
    public function update_order_status_in_sheet($order_id, $new_status) {
        try {
            // Find all rows for the order ID in the sheet
            $range = 'Sheet1!A:A'; // Assuming Order ID is in column A
            $response = $this->service->spreadsheets_values->get($this->spreadsheetId, $range);
            $values = $response->getValues();
            
            $rowsToUpdate = [];
            foreach ($values as $index => $row) {
                if ((int)$row[0] === $order_id) {
                    $rowNumber = $index + 1; // Google Sheets is 1-based index
                    $rowsToUpdate[] = 'Sheet1!F' . $rowNumber; // Assuming 'Status' column is F
                }
            }

            if (empty($rowsToUpdate)) {
                error_log('Order ID not found in Google Sheet: ' . $order_id);
                return;
            }

            // Prepare batch update for all matching rows
            $requests = [];
            foreach ($rowsToUpdate as $statusRange) {
                $requests[] = new Google_Service_Sheets_ValueRange([
                    'range' => $statusRange,
                    'values' => [[$new_status]]
                ]);
            }

            // Execute batch update
            $batchUpdateRequest = new Google_Service_Sheets_BatchUpdateValuesRequest([
                'valueInputOption' => 'RAW',
                'data' => $requests
            ]);

            $this->service->spreadsheets_values->batchUpdate($this->spreadsheetId, $batchUpdateRequest);

            error_log('Order status updated in Google Sheet for Order ID: ' . $order_id);

        } catch (Exception $e) {
            error_log('Error updating order status in sheet: ' . $e->getMessage());
        }
    }


    public function get_recent_orders_from_sheet($order_count) {
        $range = "Sheet1!A1:K";  // Adjust range as needed
        $response = $this->service->spreadsheets_values->get($this->spreadsheetId, $range);
        $values = $response->getValues();

        $orders_to_update = [];
        if (!empty($values)) {
            // Skip the header row
            for ($i = 1; $i <= $order_count && $i < count($values); $i++) {
                $row = $values[count($values) - $i];
                $order_id = $row[0];  // Adjust index based on order ID column
                $status = $row[5];    // Adjust index based on order status column
                $orders_to_update[$order_id] = $status;
            }
        }

        return $orders_to_update;
    }

}



function sheetconnect_sync_order_to_google_sheet($order_id) {
    $googleSheet = new SheetConnect_GoogleSheet();
    $googleSheet->write_order_to_sheet($order_id);
}

function sheetconnect_update_order_status_in_sheet($order_id, $old_status, $new_status, $order) {
    $googleSheet = new SheetConnect_GoogleSheet();
    $googleSheet->update_order_status_in_sheet($order_id, $new_status);
}

function sheetconnect_sync_orders_ajax_handler() {      
    if (!current_user_can('manage_options')) {
        wp_send_json_error(array('message' => 'Unauthorized user.'));
        return;
    }
    $order_count = isset($_POST['order_count']) ? intval($_POST['order_count']) : 0;
    
    if ($order_count <= 0) {
        wp_send_json_error(array('message' => 'Invalid order count.'));
        return;
    }

    try {
        $googleSheet = new SheetConnect_GoogleSheet();
        
        $orders_to_update = $googleSheet->get_recent_orders_from_sheet($order_count);

        foreach ($orders_to_update as $order_id => $status) {
            $order = wc_get_order($order_id);
            if ($order) {
                $order->set_status($status);
                $order->save();
            }
        }
        try {
            update_option('spreadsheet_last_sync_timestamp', date("d-m-Y H:i:s"));
            $data = array('status' => 'success', 'message' => 'Successfully synced '.$order_count.' orders', 'last_sync' => get_option('spreadsheet_last_sync_timestamp')); 
        } catch (Exception $e) {
            error_log('Error updating last sync timestamp: ' . $e->getMessage());
        }

        wp_send_json($data);
    } catch (Exception $e) {
        wp_send_json_error('Error syncing orders: ' . $e->getMessage());
    }
}

?>
