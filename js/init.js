function Init() {
    var menu, header, mainFrame, newHeight;
    menu = document.all.item("menu");
    for (var i = 0; i < menu.rows.length; i++) {
        for (var j = 0; j < menu.rows.item(i).cells.length; j++) {
            menu.rows.item(i).cells(j).childNodes.item(0).rows.item(2).style.position = 'absolute';
            menu.rows.item(i).cells(j).childNodes.item(0).rows.item(2).style.posTop = menu.rows.item(i).cells(j).childNodes.item(0).rows.item(1).offsetTop + menu.rows.item(i).cells(j).childNodes.item(0).rows.item(1).offsetHeight;
            newHeight = menu.rows.item(i).cells(j).childNodes.item(0).rows.item(0).offsetHeight + menu.rows.item(i).cells(j).childNodes.item(0).rows.item(1).offsetHeight + menu.rows.item(i).cells(j).childNodes.item(0).rows.item(2).offsetHeight;
            menu.rows.item(i).cells(j).height = newHeight;
        }
    }
}