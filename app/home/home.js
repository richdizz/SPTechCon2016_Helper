(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason) {
    $(document).ready(function(){
      app.initialize();
      $("#btnSignin").click(function() {
          AzureADAuth.getAccessToken()
            .then(function (token) {
                access_token = token;
                //handle token
                $.ajax({
                    url: "https://graph.microsoft.com/beta/me/joinedgroups",
                    headers: {
                        "Accept": "application/json",
                        "Authorization": "Bearer " + access_token 
                    },
                    method: "GET",
                    success: function (data) {
                        $("#view-login").hide();
                        $("#view-groups").show();

                        var html = "<ul>";
                        $(data.value).each(function(i, e) {
                           html += "<li><a href='javascript:loadGroup(\"" + e.id + "\")'>" + e.displayName + "</a></li>";
                        });
                        html += "</ul>";
                        $("#view-groups").html(html);
                    }
                });
                
            })
            .error(function (err) {
                //handle error
                $("#token").html("Error getting token");
            });
         });
      });
    };
    
    
})();

var access_token = null;
function loadGroup(id) {
    $("#view-groups").hide();
    $("#view-spinner").show();
    
    $.ajax({
        url: "https://graph.microsoft.com/v1.0/groups/" + id + "/members",
        headers: {
            "Accept": "application/json",
            "Authorization": "Bearer " + access_token 
        },
        method: "GET",
        success: function (data) {
            //Load data with members
            var html = "<ul>"
            $(data.value).each(function(i, e) {
                html += "<li>" + e.displayName + "</li>";
            });
            html += "</ul>"
            $("#view-spinner").html(html);
        }
    });
}