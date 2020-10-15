// import XlsxPopulate, { RichText } from 'xlsx-populate'

// let s = '<p>Hello World</p><ul><li>One</li><li>Two</li></ul><p>Hello World</p><ol><li>One Two</li><li><b>Three Four</b></li><li><b><i>Five </i></b><i>Six</i><ol><li><i>Seven</i></li></ol></li><li>Eight</li></ol><p><br></p>'
// let p = '<p><u>Hello</u><b>World</b></p><p><u>Hello</u><b>World</b></p>'

// XlsxPopulate.fromBlankAsync()
//     .then(workbook => {
//         const cell =  workbook.sheet(0).cell('A1')
        
//         cell.value(html2rtf(s))

//         workbook.outputAsync("base64")
//         .then(function (base64) {
//             location.href = "data:" + XlsxPopulate.MIME_TYPE + ";base64," + base64;
//         });
//     });

function html2rtf(celltext){
    let tagStyles = ['DEL', 'S', 'STRIKE', 'U', 'I', 'EM', 'B', 'STRONG', 'SUP', 'SUB', 'LI']
    let richText = new XlsxPopulate.RichText()

    if(celltext){
        let div = document.createElement('div')
        div.innerHTML = celltext.replace(/&#58;/g, ':').replace(/<br>/g, '\n').replace(/&nbsp;/, ' ')
        
        let rootChildren = div.children
        let styledElements = []
        let foundFirstText = false
        let foundFirstBlockChild = true
        let numbering = {}
        let currentLevel = 0

        let loopThroughChildNodes = (nodes) =>{
            let fragmentCount = 0
            
            for(let i=0; i<nodes.length; i++){

                let textNode = nodes[i].nodeType == 3 && nodes[i]
                
                let elementNode = nodes[i].nodeType == 1 && nodes[i]
                let elementStyles = []

                let fragmentValue = textNode ? textNode.textContent : ''
                let fragmentStyles = {}

                if(elementNode){
                    if((elementNode.nodeName == 'DIV' || elementNode.nodeName == 'P') && foundFirstBlockChild){
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
                            console.log(currentLevel)
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
                    if (!elementNode.childNodes || elementNode.childNodes.length == 0) {fragmentValue = elementNode.textContent}
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
                            case 'font-weight:bold':
                                fragmentStyles['bold'] = true; break;
                            case 'font-style:italic':
                                fragmentStyles['italic'] = true; break;
                            case 'text-decoration:underline':
                                fragmentStyles['underline'] = true; break;
                            case 'text-decoration:strikethrough':
                                fragmentStyles['strikethrough'] = true; break;
                            case 'text-decoration:line-through':
                                fragmentStyles['strikethrough'] = true; break;
                        }

                        let color = elementStyles[j].match(/^color:#?(.+)/);
                        if (color && color[1] != 'Black') fragmentStyles['fontColor'] = color[1]
                    }

                    if (styledElements.length > 0) {
                        fragmentStyles['fontSize'] = fragmentStyles['fontSize'] || 10
                        fragmentStyles['fontFamily'] = fragmentStyles['fontFamily'] || 'Arial'
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