/**
 * jQuery sdSwide
 * Version 0.02 - 05.07.2009
 * Author Stoyan Delev http://stoyandelev.com
 **/


(function($){  
	$.fn. sdswitch = function(options) {  
   
		var defaults = {  
			animateTime: 1000,
			time:1000,
			startElement: 0,
			showWindow: false,
			showTitle: false  
		};  
		
		var options = $.extend(defaults, options); 
		var el = this; 
       	
       	el.addClass('sdswitch');
		
		//first element
		var $first = el.children(':first');
		
		//active element
		var $active = el.children('.active');
		
		//if not have active element, active is last 
		if ( $active.length == 0 ) {
			$active = el.children(':eq('+ ( options.startElement - 1 ) +')');
		}
		
		//next element	
		var $next =  $active.next().length ? $active.next() : $first;
	
		$next.css({opacity: 0.0}).addClass('active').animate({opacity: 1.0}, options.animateTime, function() {
			$active.removeClass('active');
		});
		
		
		//add window
		if ( options.showWindow ) {
		
			var curNum = el.children().index($active);
			var totalChild = el.children().length;
						
			if ( $('.sdwindow').length == 0 ) {
				$('<div></div>').addClass('sdwindow').appendTo( el.parent());
			}
			
			//show title of element
			if ( options.showTitle ) {
				
				$('.sdwindow').html( $next.attr('title') );
				
			} else {
			
				$('.sdwindow').html('slide ' + (++curNum) + ' of ' +  totalChild );
			
			}
		
		}
		
		return setTimeout( function() { el.sdswitch(options); }, options.time);
		
		//var $time = options.time;
		//if ($next[0].id == $first[0].id || $active[0].id == $first[0].id)
		// 	   { $time = options.time / 2; }
		//return setTimeout( function() { el.sdswitch(options); }, $time);
	
	};  
})(jQuery)