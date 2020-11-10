// import XlsxPopulate, { RichText } from 'xlsx-populate'
// import ExcelJS from 'exceljs'

// let s = '<p>Hello World</p><ul><li>One</li><li>Two</li></ul><p>Hello World</p><ol><li>One Two</li><li><b>Three Four</b></li><li><b><i>Five </i></b><i>Six</i><ol><li><i>Seven</i></li></ol></li><li>Eight</li></ol><p><br></p>'
// let p = '<p>Hello World</p><ul><li>One</li><li>Two</li></ul><p>Hello World</p><ol><li>One Two</li><li><b>Three Four</b></li></ol><p><b>Zero</b></p><ol><li><b><i>Five </i></b><i>Six</i><ol><li><i>Seven</i></li><li>Nine</li><li>Ten</li></ol></li><li>Eight</li></ol><p><img src="https://cunning-wolf-lr1avv-dev-ed--c.documentforce.com/servlet/rtaImage?eid=0012w00000Kou4b&amp;feoid=00N2w00000FPu0u&amp;refid=0EM2w000001VIch" alt="alumini.jpg"></img></p>'
// let q = '<table class="ql-table-blob" dir="ltr" border="1" style="font-size: 11pt; font-family: Calibri; width: 0px;"><colgroup><col width="60"></col><col width="60"></col><col width="60"></col><col width="60"></col></colgroup><tbody><tr style="height: 21px;"><td colspan="1" rowspan="1" style="">One</td><td colspan="1" rowspan="1" style="">Two</td><td colspan="1" rowspan="1" style="">Three</td><td colspan="1" rowspan="1" style="">Four</td></tr><tr style="height: 21px;"><td colspan="1" rowspan="1" style="font-weight: bold;">Five</td><td colspan="1" rowspan="1" style="font-style: italic;">Six</td><td colspan="1" rowspan="1" style="font-style: italic;">Seven</td><td colspan="1" rowspan="1" style="font-weight: bold;">Eight</td></tr><tr style="height: 21px;"><td colspan="1" rowspan="1" style="">Nine</td><td colspan="1" rowspan="1" style="">Ten</td><td colspan="1" rowspan="1" style="">Eleven</td><td colspan="1" rowspan="1" style="">Twelve</td></tr></tbody></table><p><br></p>'

// XlsxPopulate.fromBlankAsync()
//     .then(workbook => {
//         const cell =  workbook.sheet(0).cell('A1')
        
//         cell.value(html2rtf(s + q + p))

//         workbook.outputAsync("base64")
//         .then(function (base64) {
//             location.href = "data:" + XlsxPopulate.MIME_TYPE + ";base64," + base64;
//         });
//     });

function html2rtf(celltext, addTableToCell=true){
    let tagStyles = ['DEL', 'S', 'STRIKE', 'U', 'I', 'EM', 'B', 'STRONG', 'SUP', 'SUB', 'LI']
    let richText = new XlsxPopulate.RichText()

    if(celltext){
        let div = document.createElement('div')
        div.innerHTML = celltext.replace(/&#58;/g, ':').replace(/<br>/g, '\n').replace(/&nbsp;/, ' ')

        if(!addTableToCell){
            let lst = div.querySelectorAll('table')
            lst.forEach(elm => {
                elm.parentNode.removeChild(elm)
            })
        }    
        
        let rootChildren = div.children
        let styledElements = []
        let foundFirstText = false
        let foundFirstBlockChild = true
        let numbering = {}
        let currentLevel = 0
        let maxLength = 0
        let tdCount = 0
        let trCount = 0
        let numCol = 0
        let currTdLen = 0
        let tableFont = {'fontFamily': 'courier'}
        let isTableChild = false

        let loopThroughChildNodes = (nodes) =>{
            let fragmentCount = 0
            
            for(let i=0; i<nodes.length; i++){

                let textNode = nodes[i].nodeType == 3 && nodes[i]
                
                let elementNode = nodes[i].nodeType == 1 && nodes[i]
                let elementStyles = []

                let fragmentValue = textNode ? textNode.textContent : ''
                let fragmentStyles = {}

                if(elementNode){
                    if((elementNode.nodeName == 'DIV' || elementNode.nodeName == 'P' || elementNode.nodeName == 'TR') && foundFirstBlockChild){
                        foundFirstBlockChild = false
                    }
                    if(elementNode.nodeName == 'UL' || elementNode.nodeName == 'OL'){
                        foundFirstBlockChild = false
                        currentLevel++
                        numbering[currentLevel] = 1
                    }
                    if (elementNode.getAttribute('style')) {
                        styledElements.push(elementNode); 
                    } else if (tagStyles.indexOf(elementNode.nodeName) > -1) {
                        styledElements.push(elementNode);
                    }
                    if (elementNode.nodeName == "LI"){
                        foundFirstBlockChild = false
                        if(elementNode.parentNode.nodeName == "UL"){
                            switch(currentLevel%2){
                                case 1:
                                    fragmentValue = " ".repeat(currentLevel*4) + "\u2022" + " " + fragmentValue; break;
                                case 0:
                                    fragmentValue = " ".repeat(currentLevel*4) + "\u25cb" + " " + fragmentValue; break;
                            }
                        }
                        else if(elementNode.parentNode.nodeName == "OL"){
                            switch(currentLevel%3){
                                case 0:
                                    fragmentValue = " ".repeat(currentLevel*4) + romanize(numbering[currentLevel]) + ". " + fragmentValue; break;
                                case 1:
                                    fragmentValue = " ".repeat(currentLevel*4) + numbering[currentLevel] + ". " + fragmentValue; break;
                                case 2:
                                    fragmentValue = " ".repeat(currentLevel*4) + colName(numbering[currentLevel]-1) + ". " + fragmentValue; break;
                            }
                        }
                        numbering[currentLevel]++
                    }
                    if(elementNode.nodeName == 'IMG'){
                        fragmentValue = "<" + elementNode.getAttribute('alt') + ">"
                        fragmentStyles["italic"] = true
                    }
                    if(elementNode.nodeName == 'TABLE'){
                        isTableChild = true
                        let lst = textNodesUnder(elementNode)
                        maxLength = Math.max(...lst.map(e => e.length))
                        numCol = elementNode.getElementsByTagName('tr')[0].childElementCount 

                    }
                    if(elementNode.nodeName == "TR"){
                        trCount++
                        if(trCount === 1){
                            let dash = ((maxLength + 4) * numCol) + numCol + 1
                            fragmentValue = "-".repeat(dash) + "\n"
                        }
                    }
                    if(elementNode.nodeName == "TD"){
                        tdCount++
                        if(tdCount === 1){
                            fragmentValue = "|  "
                        }
                        else{
                            fragmentValue += "  "
                        }
                        currTdLen = elementNode.textContent.length
                    }
                    //if (!elementNode.childNodes || elementNode.childNodes.length == 0) {fragmentValue = elementNode.textContent}
                }
                if(fragmentValue){
                    let checkTagStyles = (_node) => {
                        if (tagStyles.indexOf(_node.nodeName) > -1) {
                            if (_node.nodeName == 'DEL' || _node.nodeName == 'S' || _node.nodeName == 'STRIKE')
                                fragmentStyles['strikethrough'] = true;
                            if (_node.nodeName == 'U')
                                fragmentStyles['underline'] = true;
                            if (_node.nodeName == 'I' || _node.nodeName == 'EM')
                                fragmentStyles['italic'] = true;
                            if (_node.nodeName == 'B' || _node.nodeName == 'STRONG')
                                fragmentStyles['bold'] = true;
                            if (_node.nodeName == 'SUP')
                                fragmentStyles['superscript'] = true;
                            if (_node.nodeName == 'SUB')
                                fragmentStyles['superscript'] = true;
                        }
                    }

                    if (foundFirstText && !foundFirstBlockChild) {
                        fragmentValue = '\n' + fragmentValue;
                    }

                    for (let j = 0; j < styledElements.length; j++) {
                        let isParent = styledElements[j].contains(elementNode || textNode);
                        if (isParent) {
                            elementStyles.push(styledElements[j].getAttribute('style'));
                            checkTagStyles(styledElements[j])
                        }
                    }

                    if (elementNode) {
                        if (elementNode.getAttribute('style')) {
                            elementStyles.push(elementNode.getAttribute('style'));
                        }
                        checkTagStyles(elementNode)
                    }

                    elementStyles = elementStyles.join(';').split(';')

                    for (let j = 0; j<elementStyles.length;j++) {
                        // individual styles
                        switch (elementStyles[j]) {
                            case 'font-weight: bold':
                                fragmentStyles['bold'] = true; break;
                            case 'font-style: italic':
                                fragmentStyles['italic'] = true; break;
                            case 'text-decoration: underline':
                                fragmentStyles['underline'] = true; break;
                            case 'text-decoration: strikethrough':
                                fragmentStyles['strikethrough'] = true; break;
                            case 'text-decoration: line-through':
                                fragmentStyles['strikethrough'] = true; break;
                        }

                        let color = elementStyles[j].match(/^color:#?(.+)/);
                        if (color && color[1] != 'Black') fragmentStyles['fontColor'] = color[1]
                    }

                    // if (styledElements.length > 0) {
                    //     fragmentStyles['fontSize'] = fragmentStyles['fontSize'] || 10
                    //     fragmentStyles['fontFamily'] = fragmentStyles['fontFamily'] || 'Arial'
                    // }

                    if(isTableChild){
                        fragmentStyles['fontFamily'] = 'courier'
                    }
                    richText.add(fragmentValue, fragmentStyles);
                
                    if (/(?:^\s)|(?:\s$)/.test(fragmentValue)) {
                        richText.get(fragmentCount)._valueNode.attributes['xml:space'] = 'preserve'
                    }

                    fragmentCount += 1;

                    if (foundFirstText == false) {
                        foundFirstText = true
                    }
                    foundFirstBlockChild = true
                }

                if (elementNode.childNodes) {loopThroughChildNodes(elementNode.childNodes)}

                if(elementNode.nodeName == 'OL' || elementNode.nodeName == 'UL'){
                    numbering[currentLevel] = 0
                    currentLevel--
                }
                if(elementNode.nodeName == "TABLE"){
                    isTableChild = false
                }
                if(elementNode.nodeName == "TD"){
                    if(numCol === tdCount){
                        richText.add(" ".repeat(maxLength - currTdLen) + "  |" + "\n", tableFont);
                    }
                    else{
                        richText.add(" ".repeat(maxLength - currTdLen) + "  |", tableFont);
                    }
                }
                if(elementNode.nodeName == "TR"){
                    let dash = ((maxLength + 4) * numCol) + numCol + 1
                    let str = "-".repeat(dash)
                    richText.add(str, tableFont)
                    tdCount = 0
                }
            }
            
        }

        loopThroughChildNodes(rootChildren)

        return richText
    }

    return richText
}

function colName(n) {
    var ordA = 'a'.charCodeAt(0);
    var ordZ = 'z'.charCodeAt(0);
    var len = ordZ - ordA + 1;
  
    var s = "";
    while(n >= 0) {
        s = String.fromCharCode(n % len + ordA) + s;
        n = Math.floor(n / len)-1;
    }
    return s;
}

function romanize(num) {
    var lookup = {M:1000,CM:900,D:500,CD:400,C:100,XC:90,L:50,XL:40,X:10,IX:9,V:5,IV:4,I:1},roman = '',i;
    for ( i in lookup ) {
      while ( num >= lookup[i] ) {
        roman += i;
        num -= lookup[i];
      }
    }
    return roman; 
}

function textNodesUnder(el){
    var n, a=[], walk=document.createNodeIterator(el,NodeFilter.SHOW_TEXT,null,false);
    while(n=walk.nextNode()) a.push(n.textContent);
    return a;
}