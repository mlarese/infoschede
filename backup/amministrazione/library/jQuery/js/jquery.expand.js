////////////////////////////
// http://adipalaz.awardspace.com/experiments/jquery/expand.html
// * When using this script, please keep the above url intact.
///////////////////////////
(function($) {
$.fn.expandAll = function(options) {
    var defaults = {
         expTxt : '[Expand All]',
         cllpsTxt : '[Collapse All]',
         cllpsEl : 'div.collapse', // the collapsible element
         trigger : '.expand', // the element that triggers the toggle action
         ref : '.expand', // the switch 'Expand All/Collapse All' is inserted before the first 'ref'
         showMethod : 'show',
         hideMethod : 'hide',
         state : 'hidden', // the collapsible elements are hidden by default
         speed : 0
    };
    var o = $.extend({}, defaults, options);
    
    var container = '#' + $(this).attr('id'),
        toggleTxt = o.expTxt;
    if (o.state == 'hidden') {$(this).find(o.cllpsEl).hide();} else {toggleTxt = o.cllpsTxt;};
    $(container + ' '+ o.ref + ':first').before('<p id="switch"><a href="#">' + toggleTxt + '</a></p>');     
    return this.each(function() {
        $(this).find('#switch a').click(function() {
        var $cllps = $(this).closest(container).find(o.cllpsEl),
            $tr = $(this).closest(container).find(o.trigger);
        if ($(this).text() == o.expTxt) {
          $(this).text(o.cllpsTxt);
          $tr.addClass('open');
          $cllps[o.showMethod](o.speed);
        } else {
          $(this).text(o.expTxt);
          $tr.removeClass('open');
          $cllps[o.hideMethod](o.speed);
        }
        return false;
    });
});};
$.fn.toggler = function(options) {
    var defaults = {
         cllpsEl : 'div.collapse',
         method : 'slideToggle',
         speed : 'slow'
    };
    var o = $.extend({}, defaults, options);
    
    $(this).wrapInner('<a style="display:block" href="#" title="Expand/Collapse" />');
    var container = $(this).attr('id');   
    return this.each(function() {
    $(this).click(function() {
        $(this).toggleClass('open')
        .next(o.cllpsEl)[o.method](o.speed);
        return false;
    });
});};
//http://www.learningjquery.com/2008/02/simple-effects-plugins:
$.fn.fadeToggle = function(speed, easing, callback) {
    return this.animate({opacity: 'toggle'}, speed, easing, callback);
};
$.fn.slideFadeToggle = function(speed, easing, callback) {
    return this.animate({opacity: 'toggle', height: 'toggle'}, speed, easing, callback);
};
})(jQuery);