var iframeUpload = {
    init: function() {
        jQuery('body').append('<iframe class="hidden" name="uploadiframe" onload="iframeUpload.complete();"></iframe>');
        jQuery('form').prop('target','uploadiframe');
        jQuery('form').on('submit',iframeUpload.started);
    },
    started: function() {
        jQuery('#response').removeClass().addClass('loading').html('Loading, please wait.').show();
        //jQuery('form.').hide()
    },
    complete: function(){
        jQuery('form').show();
        var response = jQuery("iframe").contents().text();
        if(response){
            response = jQuery.parseJSON(response);
            jQuery('#response').removeClass()
            .addClass((response.status == 1) ? alert('File uploaded successfully') : alert('Failed to upload the file'))
            .html(response.message);
        }
    }
};