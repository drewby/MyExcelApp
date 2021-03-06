(function(){
  'use strict';
   var fill = d3.scale.category20();
  
  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();
      jQuery('#get-data-from-selection').click(getDataFromSelection);
    });
  };

  // Reads data from current document selection and displays a notification
  function getDataFromSelection(){
    if (Office.context.document.getSelectedDataAsync) {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
        function(result){
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            app.showNotification('The selected text is:', '"' + result.value + '"');
            
              var words = result.value.split('\n').map(function(d) {
                var word = d.split('\t');
                return {text: word[0], size: word[1]}
              });
            
              d3.layout.cloud().size([400, 400])
                .words(words)
                .padding(5)
                .rotate(function() { return ~~(Math.random() * 2) * 90; })
                .font("Impact")
                .fontSize(function(d) { return d.size / 2; })
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
    document.getElementById("content-svg").innerHTML = "";
    d3.select("#content-svg").append("svg")
        .attr("width", 400)
        .attr("height", 400)
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
