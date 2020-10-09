import XlsxPopulate, { RichText } from 'xlsx-populate'

// let s = <ul><li>One</li><li>Two</li></ul>

XlsxPopulate.fromBlankAsync()
    .then(workbook => {
        const cell =  workbook.sheet(0).cell('A1')
        
        cell.value(html2rtf(s))


        workbook.outputAsync("base64")
        .then(function (base64) {
            location.href = "data:" + XlsxPopulate.MIME_TYPE + ";base64," + base64;
        });
    });

function html2rtf(celltext){
    let tagStyles = ['DEL', 'S', 'STRIKE', 'U', 'I', 'EM', 'B', 'STRONG', 'SUP', 'SUB']
    let richText = new XlsxPopulate.RichText()

    if(celltext){
        let div = document.createElement('div')
        div.innerHTML = celltext.replace(/&#58;/g, ':').replace(/<br>/g, '\n').replace(/&nbsp;/, ' ')
        
        let rootChildren = div.children
        let styledElements = []
        let foundFirstText = false
        let foundFirstBlockChild = true

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
                    if (elementNode.getAttribute('style')) {
                        styledElements.push(elementNode); 
                    } else if (tagStyles.indexOf(elementNode.nodeName) > -1) {
                        styledElements.push(elementNode);
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
                        foundFirstBlockChild = true
                    }
                }

                if (elementNode.childNodes) {loopThroughChildNodes(elementNode.childNodes)}
            }
            
        }

        for (let i = 0; i < rootChildren.length; i++) {
            loopThroughChildNodes(rootChildren[i].childNodes)
        }

        return richText
    }

    return richText
}
