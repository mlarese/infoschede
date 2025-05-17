
	function resizeImage(id_container) {
		resizeImageWOption(id_container, false);
	}
	
	// Actual resize function
    function resizeImageWOption(id_container, toCenter) {
		// Actual resize function
		//function resizeImage(id_container) {
		//var toCenter;
		//toCenter = false;
		
		
		//"z-index":"-1",
		// "overflow":"hidden",
        $("#"+id_container).css({
            "position":"fixed",
            "top":"0px",
            "left":"0px",
            "width":$(window).width() + "px",
            "height":$(window).height() + "px"
        });
		
		// Image relative to its container
		//$("#"+id_container).children('img').css("position", "relative");

        // Resize the img object to the proper ratio of the window.
        var iw = $("#"+id_container).children('img').width();
        var ih = $("#"+id_container).children('img').height();
       
        if ($(window).width() > $(window).height()) {
            //console.log(iw, ih);
            if (iw >= ih) {
                var fRatio = iw/ih;
                $("#"+id_container).children('img').css("width",$(window).width() + "px");
                $("#"+id_container).children('img').css("height",Math.round($(window).width() * (1/fRatio)));

                var newIh = Math.round($(window).width() * (1/fRatio));

                if(newIh < $(window).height()) {
                    var fRatio = ih/iw;
                    $("#"+id_container).children('img').css("height",$(window).height());
                    $("#"+id_container).children('img').css("width",Math.round($(window).height() * (1/fRatio)));
                }
            } else {
                var fRatio = ih/iw;
                $("#"+id_container).children('img').css("height",$(window).height());
                $("#"+id_container).children('img').css("width",Math.round($(window).height() * (1/fRatio)));
            }
        } else {
            var fRatio = ih/iw;
            $("#"+id_container).children('img').css("height",$(window).height());
            $("#"+id_container).children('img').css("width",Math.round($(window).height() * (1/fRatio)));
        }
		
		// Center the image
		if (typeof(toCenter) == 'undefined' || toCenter) {
			if ($("#"+id_container).children('img').width() > $(window).width()) {
				var this_left = ($("#"+id_container).children('img').width() - $(window).width()) / 2;
				$("#"+id_container).children('img').css({
					"top"  : 0,
					"left" : -this_left
				});
			}
			if ($("#"+id_container).children('img').height() > $(window).height()) {
				var this_height = ($("#"+id_container).children('img').height() - $(window).height()) / 2;
				$("#"+id_container).children('img').css({
					"left" : 0,
					"top" : -this_height
				});
			}
		}

        $("#"+id_container).css({
			"visibility" : "visible"
		});

		// Allow scrolling again
		$("body").css({
            "overflow":"auto"
        });
		
        
    }
