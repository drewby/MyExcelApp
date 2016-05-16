(function(){
  'use strict';
   var fill = d3.scale.category20();
  
  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();
      jQuery('#get-data-from-selection').click(getDataFromSelection);
      jQuery('#content-main').append('div').innerHTML = "Hello world";
    });
  };

  // Reads data from current document selection and displays a notification
  function getDataFromSelection(){
    if (Office.context.document.getSelectedDataAsync) {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
        function(result){
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            app.showNotification('The selected text is:', '"' + result.value + '"');
            
              d3.layout.cloud().size([300, 300])
                .words([
                  "Hello", "world", "normally", "you", "want", "more", "words",
                  "than", "this"].map(function(d) {
                  return {text: d, size: 10 + Math.random() * 90};
                }))
                .padding(5)
                .rotate(function() { return ~~(Math.random() * 2) * 90; })
                .font("Impact")
                .fontSize(function(d) { return d.size; })
                .on("end", draw)
                .start();
          } else {
            app.showNotification('Error:', result.error.message);
          }
        }
        );
    } else {
      app.showNotification('Error:', 'Reading selection data not supported by host application.');
    }
  }
  
  function draw(words) {
    d3.select("#content-main").append("svg")
        .attr("width", 300)
        .attr("height", 300)
      .append("g")
        .attr("transform", "translate(150,150)")
      .selectAll("text")
        .data(words)
      .enter().append("text")
        .style("font-size", function(d) { return d.size + "px"; })
        .style("font-family", "Impact")
        .style("fill", function(d, i) { return fill(i); })
        .attr("text-anchor", "middle")
        .attr("transform", function(d) {
          return "translate(" + [d.x, d.y] + ")rotate(" + d.rotate + ")";
        })
        .text(function(d) { return d.text; });
  }
})();
