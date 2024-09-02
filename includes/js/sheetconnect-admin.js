// js/sheetconnect-admin.js

jQuery(document).ready(function($) {
    $('#sync-orders-button').on('click', function(e) {
        e.preventDefault();

        var orderCount = $('#order-count-input').val();

        if (orderCount) {
            // Trigger AJAX request
            $.ajax({
                url: sheetconnect_ajax.ajax_url,
                type: 'POST',
                data: {
                    action: 'sheetconnect_sync_orders',
                    order_count: orderCount,
                },
                beforeSend: function() {
                    // Optionally show a loading spinner or disable the button
                    $('#sync-orders-button').prop('disabled', true).text('Syncing...');
                },
                success: function(response) {
                        // console.log(response);
                        if(response.status === 'success') {
                            alert(response.message);
                            $('#sync-orders-button').prop('disabled', false).text('Sync Orders');
                            $('#last-sync').text('Last Sync: ' + response.last_sync);
                        }
                },
                error: function(xhr, status, error) {
                    alert('An error occurred: ' + error.data);
                    $('#sync-orders-button').prop('disabled', false).text('Sync Orders');
                }
            });
        } else {
            alert('Please enter a valid number of orders to sync.');
        }
    });
});
