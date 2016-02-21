(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason) {
    $(document).ready(function(){
      app.initialize();
      $("#btnSignin").click(function() {
          AzureADAuth.getAccessToken()
            .then(function (token) {
                //handle token
                $("#token").html(token);
            })
            .error(function (err) {
                //handle error
                $("#token").html("Error getting token");
            });
         });
      });
    };
})();
