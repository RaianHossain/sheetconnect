<?php
function sheetconnect_add_admin_menu() {
    add_submenu_page(
        'woocommerce', // Parent slug - WooCommerce menu
        'SheetConnect', // Page title
        'SheetConnect', // Menu title
        'manage_options', // Capability
        'sheetconnect', // Menu slug
        'sheetconnect_settings_page' // Callback function
    );
}
add_action('admin_menu', 'sheetconnect_add_admin_menu');


// Render the settings page
function sheetconnect_enqueue_bootstrap() {
    wp_enqueue_script('jquery');
    // Enqueue Bootstrap 5 CSS
    wp_enqueue_style('bootstrap-css', 'https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css', array(), '5.3.2');
    
    // Enqueue Bootstrap 5 JS Bundle (includes Popper)
    wp_enqueue_script('bootstrap-js', 'https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js', array('jquery'), '5.3.2', true);

    // Enqueue your custom JavaScript for handling the AJAX
    wp_enqueue_script('sheetconnect-admin-js', plugin_dir_url(__FILE__) . 'js/sheetconnect-admin.js', array('jquery'), '1.0.0', true);
    
    // Localize script to pass AJAX URL to JavaScript
    wp_localize_script('sheetconnect-admin-js', 'sheetconnect_ajax', array(
        'ajax_url' => admin_url('admin-ajax.php'),
        'nonce' => wp_create_nonce('sheetconnect_sync_nonce'), // Secure nonce for verification
    ));
}
add_action('admin_enqueue_scripts', 'sheetconnect_enqueue_bootstrap');


function sheetconnect_settings_page() {
    $spreadsheet_id = get_option('sheetconnect_google_spreadsheet_id', '');
    $service_account_file_name = get_option('sheetconnect_service_account_file_name', '');
    $sheet_connected = !empty($spreadsheet_id) && !empty($service_account_file_name);
    $last_sync = get_option('spreadsheet_last_sync_timestamp', "");

    if ($_SERVER['REQUEST_METHOD'] === 'POST') {
        error_log("spreadsheet id: " . $_POST['sheetconnect_google_spreadsheet_id']);
        // Save Spreadsheet ID        
        if (isset($_POST['sheetconnect_google_spreadsheet_id'])) {
            update_option('sheetconnect_google_spreadsheet_id', sanitize_text_field($_POST['sheetconnect_google_spreadsheet_id']));
            $spreadsheet_id = sanitize_text_field($_POST['sheetconnect_google_spreadsheet_id']);
            // error_log("spreadsheet id: " . $_POST['sheetconnect_google_spreadsheet_id']);
            if(empty($_POST['sheetconnect_google_spreadsheet_id'])) {
              $sheet_connected = false;
            } else {
              $sheet_connected = true;
            }
        }

        if (isset($_FILES['sheetconnect_service_account_file']) && $_FILES['sheetconnect_service_account_file']['type'] === 'application/json') {
            $uploaded_file = $_FILES['sheetconnect_service_account_file'];
            $upload_dir = wp_upload_dir()['basedir'] . '/sheetconnect';
            if (!file_exists($upload_dir)) {
                mkdir($upload_dir, 0755, true);
            }
            $file_path = $upload_dir . '/' . basename($uploaded_file['name']);
            if (move_uploaded_file($uploaded_file['tmp_name'], $file_path)) {
                update_option('sheetconnect_service_account_file_name', basename($uploaded_file['name']));
                $service_account_file_name = basename($uploaded_file['name']);
            }
            error_log("service account file");
        }
    }
    ?>
    <div class="wrap">
           <h1>SheetConnect Settings</h1>

           <p>This plugin connects WooCommerce orders with a Google Spreadsheet. Every order data is populated to the sheet, allowing multiple people to work on it collaboratively. If someone changes the order status from the sheet, they can sync the changes back to WooCommerce.</p>

            <button type="button" class="btn btn-outline-warning" data-bs-toggle="modal" data-bs-target="#exampleModal">
              How To Configure
            </button>

            <!-- Modal -->
            <div class="modal fade" id="exampleModal" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
              <div class="modal-dialog">
                <div class="modal-content">
                  <div class="modal-header">
                    <h5 class="modal-title" id="exampleModalLabel">Configuration Guidancee</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                  </div>
                  <div class="modal-body">
                    
                   <p>Step-by-step instructions to configure the Google Sheet integration:</p>
                   <ol>
                       <li>Enter your Google Spreadsheet ID in the Configuration tab.</li>
                       <li>Upload your Google Service Account JSON key file.</li>
                       <li>Click "Save Changes" to connect your spreadsheet.</li>
                   </ol>
                  </div>
                  <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                    <!-- <button type="button" class="btn btn-primary">Save changes</button> -->
                  </div>
                </div>
              </div>
            </div>
           
           <div class="mt-4">
               <ul class="nav nav-tabs" id="myTab" role="tablist">
                 <li class="nav-item" role="presentation">
                   <button class="nav-link active" id="home-tab" data-bs-toggle="tab" data-bs-target="#home" type="button" role="tab" aria-controls="home" aria-selected="true">Configuration</button>
                 </li>
                 <li class="nav-item" role="presentation">
                   <button class="nav-link" id="profile-tab" data-bs-toggle="tab" data-bs-target="#profile" type="button" role="tab" aria-controls="profile" aria-selected="false">Synchronization</button>
                 </li>
                 <!-- <li class="nav-item" role="presentation">
                   <button class="nav-link" id="contact-tab" data-bs-toggle="tab" data-bs-target="#contact" type="button" role="tab" aria-controls="contact" aria-selected="false">Contact</button>
                 </li> -->
              </ul>
               <div class="tab-content" id="myTabContent">
                 <div class="tab-pane fade show active" id="home" role="tabpanel" aria-labelledby="home-tab">
                     <div class="p-4">
                       <form method="post" enctype="multipart/form-data" action="">
                          <div class="row g-3 align-items-center m-3">
                             <div class="col-auto">
                               <label for="inputPassword6" class="col-form-label">Spreadsheet ID:</label>
                             </div>
                             <div class="col-auto ms-2">
                               <input type="text" name="sheetconnect_google_spreadsheet_id" id="inputPassword6" class="form-control" aria-describedby="passwordHelpInline"
                               style="width: 450px"
                               value="<?php echo $spreadsheet_id ?>"
                               >
                             </div>
                             <!-- <div class="col-auto">
                               <span id="passwordHelpInline" class="form-text">
                                 Ex: https://docs.google.com/spreadsheets/d/{Spredsheet Id}/edit?gid=0#gid=0
                               </span>
                             </div> -->
                          </div>

                          <div class="row g-3 align-items-center m-3">
                             <div class="col-auto">
                               <label for="inputPassword6" class="col-form-label">Google Key Json:</label>
                             </div>
                             <div class="col-auto">
                               <div class="input-group">
                                <input type="file" name="sheetconnect_service_account_file" class="form-control" id="inputGroupFile04" aria-describedby="inputGroupFileAddon04" aria-label="Upload" style="width: 450px" />  
                                                            
                              </div>
                             </div>
                             <?php if ($service_account_file_name != ""): ?>
                             <div class="col-auto">
                               <span id="passwordHelpInline" class="form-text">
                                 File Uploaded
                               </span>
                             </div>
                             <?php endif ?>
                           </div>
                          
                          <button type="submit" class="btn btn-primary m-3 ms-4">Save Change</button>
                       </form>
                     </div>
                 </div>
                 <div class="tab-pane fade" id="profile" role="tabpanel" aria-labelledby="profile-tab">
                    <div class="p-4">
                      <form id="sync-orders-form">
                        <div class="row g-3 align-items-center p-4">
                            <div class="col-auto">
                              <label for="inputPassword6" class="col-form-label">Number of rows:</label>
                            </div>
                            <div class="col-auto">
                              <input id="order-count-input" name="order_count" type="number" id="inputPassword6" class="form-control" aria-describedby="passwordHelpInline">
                            </div>
                            <div class="col-auto">
                              <span id="passwordHelpInline" class="form-text">
                                <button type="button" id="sync-orders-button" class="btn btn-primary">Synch With Sheet</button>
                              </span>
                            </div>
                            <p><i id="last-sync">Last Sync: <?php echo $last_sync ?></i></p>
                        </form>
                        </div>
                    </div>
                 </div>
                 <!-- <div class="tab-pane fade" id="contact" role="tabpanel" aria-labelledby="contact-tab">Contact</div> -->
               </div>              
           </div>             
           

           <hr>
           
          <?php $sheet_url = "https://docs.google.com/spreadsheets/d/".$spreadsheet_id; ?>  
           <div class="sheetconnect-spreadsheet-table">
               <h4>Spreadsheet Details</h4>
               <?php if ($sheet_connected): ?>
                   <table class="table table-bordered mt-4">
                     <tbody>
                        <tr>
                           <td>Your Spreadsheet: <?php echo $spreadsheet_id ?></td>
                           <td class="text-center">
                              <!-- <button type="button" class="btn btn-outline-info" 
                              >Go to Spreadsheet</button> -->
                              <a href="<?php echo esc_url($sheet_url); ?>" target="_blank" class="button">Open Spreadsheet</a>
                           </td>
                        </tr>
                      </tbody>
                   </table>
               <?php else: ?>
                   <p>No spreadsheet connected. Please configure the settings above.</p>
               <?php endif; ?>
           </div>          
    </div>
    <?php
}
?>
