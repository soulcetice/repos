var heatmapchart, grid;

function draw() {
    heatmapchart = new Tee.Chart("canvas");

    grid = heatmapchart.addSeries(new Tee.ColorGrid());

    grid.horizAxis = "both";
    grid.vertAxis = "both";

    addSampleData(0);

    heatmapchart.title.text = "Heatmap";

    heatmapchart.panel.transparent = true;
    heatmapchart.axes.left.grid.centered = true;
    heatmapchart.axes.bottom.grid.centered = true;

    heatmapchart.axes.left.labels.fixedDecimals = true;
    heatmapchart.axes.left.labels.decimals = 2;
    heatmapchart.axes.right.labels.fixedDecimals = true;
    heatmapchart.axes.right.labels.decimals = 2;

    heatmapchart.tools.add(new Tee.ToolTip(heatmapchart));

    heatmapchart.draw();
}

function addSampleData(index) {

    grid.data.values = [
        [23, 15, 12, 8, 39, 50, 34], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27],
        [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27],
        [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27],
        [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27],
        [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27],
        [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27],
        [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27], [19, 7, 31, 23, 12, 40, 27],
        [19, 7, 31, 23, 12, 40, 27], [1, 26, 18, 39, 20, 6, 11]
    ];

    grid.dataChanged = true;
    grid.decimals = 2;

    grid.chart.zoom.reset();
}

$(function () {
    $('#add').hover(function () {
        //push
        for (k = 0; k < 5; k++) {
            for (i = 0; i < 37; i++) {
                heatmapchart.series.items[0].data.values[i].push(Math.random() * 100);
            }
            //splice
            if (heatmapchart.series.items[0].data.values[0].length > 500) {
                for (i = 0; i < 37; i++) {
                    heatmapchart.series.items[0].data.values[i].splice(0, 1);
                }
            }
            //min max
            grid.chart.axes.top.setMinMax(heatmapchart.series.items[0].data.values[0].length - 500, heatmapchart.series.items[0].data.values[0].length - 1);
            grid.chart.axes.bottom.setMinMax(heatmapchart.series.items[0].data.values[0].length - 500, heatmapchart.series.items[0].data.values[0].length - 1);
            //redraw
            grid.dataChanged = true;
            heatmapchart.draw();
        }
    }
    );

    draw();
});        