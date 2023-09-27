 // dimension of  excel sheet-------- 
 let column=26;
 let Row=100;
 let currentCell;
 let prevcell;
 let cutCell;
 let lastBtn;
 let matrix = new Array(Row)
 let numSheet=1;
 let currentSheet=1;
 let prevSheet=1;
 let ArrMatrix='ArrMatrix';
 //storage for all sheet cell content
 function createNewMatrix(){
    for(let row=0;row<Row;row++)
    {
       matrix[row]=new Array(column);
       for(let col=0;col<column;col++)
       {
           matrix[row][col]={};
       }
    }
 }
 //creating matrix for the first time
 createNewMatrix();

 //assecing element---------------
const TheadRow =document.getElementById("TheadRow");
const Tbody = document.getElementById("tbody");
const currentCellHeading=document.querySelector(".currentCell")
const currentCellText = document.getElementById("currentCellText");
const buttonContainer=document.querySelector(".footer")


// bold italic  and underline btn 
 const boldbtn = document.getElementById("boldbtn")
 const italicbtn = document.getElementById("italicbtn")
 const underlinebtn = document.getElementById("underlinebtn")
 const copyBtn=document.getElementById("copy")
 const cutBtn=document.getElementById("cut")
 const pasteBtn=document.getElementById("paste")
 const DownloadBTn=document.getElementById("download")
 const uploadbtn=document.getElementById("upload-input")
 const addSheetBTn=document.getElementById("add-sheet")

// alignment elements -------
const alignLeft = document.getElementById("align-left") ;
const alignCenter = document.getElementById("align-center");
const alignRight = document.getElementById("align-right");

//Dropdown elements elements -------
const fontsize = document.getElementById("font-size");
const fontStyle = document.getElementById("font-style");

// text color and bg color input element 
const textcolor = document.getElementById("textColor")
const bgcolor = document.getElementById("bgcolor")

//   cloumn  genrator function-----------------
function colGen(typeOfcell,tableRow,isInnerText,rowNo)
{
   for(let col=0;col<column;col++)
   {
    const cell = document.createElement(typeOfcell);
    if(isInnerText){
        cell.innerText = String.fromCharCode(col + 65);
        cell.setAttribute("id",String.fromCharCode(col + 65));
    }
    else{
        cell.setAttribute("id",`${String.fromCharCode(col + 65)}${rowNo}`);
        cell.setAttribute("contenteditable",true);
        cell.addEventListener("input",updateObectinMatrix)
        cell.addEventListener("focus",(event)=>focusCurrentElement(event.target))
    }
    tableRow.append(cell);
   }
}
colGen("th",TheadRow,true)

// getting current cell and setting color of respective header
 function focusCurrentElement(cell){
    currentCell=cell;
    if(prevcell){

        setHeaderColor(prevcell.id[0],prevcell.id.substring(1),"")
    }
    currentCellHeading.innerText=`${cell.id}`
    currentCell.addEventListener("keyup",()=>{
        currentCellText.value=currentCell.innerText;
    })
    currentCellText.addEventListener("keyup",()=>{
        currentCell.innerText=currentCellText.value
    })
    setHeaderColor(cell.id[0],cell.id.substring(1),"rgba(106, 255, 77, 0.459)");
    prevcell=currentCell;
 }

 // function for setting color of header
 function setHeaderColor(colID,rowID,color)
 {

    const colH = document.getElementById(colID);
    const rowH = document.getElementById(rowID);
    colH.style.backgroundColor = color;
    rowH.style.backgroundColor = color;

 }

 // genreating row.........
 function TbodyGen(){
    Tbody.innerHTML='';
    for(let row=1;row<=Row;row++)
    {
        const tr = document.createElement("tr");
        const th = document.createElement("th");
        th.innerText=row;
        th.setAttribute("id",row)
        tr.append(th);
        colGen("td",tr,false,row);
        Tbody.append(tr);
    }
 }
 // generating tbody for first time
TbodyGen();

// function for updating sheet content in the matrix 
function updateObectinMatrix() {

    let id = currentCell.id
    let col = id[0].charCodeAt(0)-65;
    let row = id.substring(1)-1;
    matrix[row][col]={
        text : currentCell.innerText,
        style : currentCell.style.cssText,
        id : id,
    };
}
 
// download matrix/sheet content
DownloadBTn.addEventListener("click",()=>{
        // 2d matrix into a memory that's accessible outside
        const matrixString = JSON.stringify(matrix);
        // matrixString -> into a blob
        const blob = new Blob([matrixString],{ type: 'application/json'});
        console.log(blob);
        const link = document.createElement('a');
        // createObjectURL converts my blob to link
        link.href = URL.createObjectURL(blob);
        link.download = 'table.json';
        link.click();
})
// upload file
uploadbtn.addEventListener("input",(event)=>{
    const file = event.target.files[0];
    // FileReader helps me to read my blod
    if(file){
      const reader = new FileReader();
      reader.readAsText(file);
      // readAsText will trigger onload method
      // of reader instance
      reader.onload = function(event){
        const fileContent=JSON.parse(event.target.result);
        matrix=fileContent;
        renderMatrix()
      }
    }
})
   

//   giving font style to current cell text
 fontStyle.addEventListener("change",()=>{
    currentCell.style.fontFamily = fontStyle.value;
    updateObectinMatrix()
 })

 //giving fontsize to current cell text
fontsize.addEventListener("change",()=>{
 currentCell.style.fontSize= fontsize.value;
 updateObectinMatrix()
})

// aligning text of current cell -----------
function alignItems(element,value){
    element.addEventListener("click",()=>{
        currentCell.style.textAlign=value;
       updateObectinMatrix()
    })  
}

alignItems(alignCenter,"center");
alignItems(alignRight,"right");
alignItems(alignLeft,"left");

// giving  text ang bg color to current cell
textcolor.addEventListener("input",()=>{
    currentCell.style.color = textcolor.value;
    updateObectinMatrix()
})

bgcolor.addEventListener("input",()=>{
    currentCell.style.backgroundColor = bgcolor.value;
    updateObectinMatrix()
})

// bold   italic underlin current cell text 
boldbtn.addEventListener("click",()=>{
    if(currentCell.style.fontWeight==="bold")
    {
        currentCell.style.fontWeight="normal";
        boldbtn.style.backgroundColor="";
    }
    else{
        currentCell.style.fontWeight="bold";
        boldbtn.style.backgroundColor="darkgrey";
    }
    updateObectinMatrix()
})
italicbtn.addEventListener("click",()=>{
    if(currentCell.style.fontStyle ==="italic")
    {
        currentCell.style.fontStyle="normal";
        italicbtn.style.backgroundColor="";
    }
    else{
        currentCell.style.fontStyle="italic";
        italicbtn.style.backgroundColor="darkgrey";
    }
    updateObectinMatrix()
})
underlinebtn.addEventListener("click",()=>{
    if(currentCell.style.textDecoration==="underline")
    {
        console.log(currentCell)
        currentCell.style.textDecoration="none";
        underlinebtn.style.backgroundColor="";
    }
    else{
        currentCell.style.textDecoration="underline";
        underlinebtn.style.backgroundColor="darkgrey";
    }
    updateObectinMatrix()
})

// copy cut and paste of current cell text

copyBtn.addEventListener("click",()=>{
    lastBtn='copy'
    console.log("het")
    cutCell={
        text: currentCell.innerText,
        style: currentCell.style.cssText,
    } 
})

cutBtn.addEventListener("click",()=>{
    lastBtn='cut'
    cutCell={
        text: currentCell.innerText,
        style: currentCell.style.cssText,
    } 
    currentCell.innerText=''
    currentCell.style.cssText=''
    updateObectinMatrix()
})

pasteBtn.addEventListener("click",()=>{
    currentCell.innerText=cutCell.text
    currentCell.style.cssText=cutCell.style
    if(lastBtn ==='cut')
    {
        cutCell = undefined;
    }
    updateObectinMatrix()
})

// Addind new sheets ..........
 function generateNextSheetBtn(){

    const sheetDiv= document.createElement("div");
    numSheet++;
    prevSheet=numSheet;
    currentSheet=numSheet;
    sheetDiv.innerText=`Sheet ${currentSheet}`;
    sheetDiv.setAttribute('id',`sheet-${currentSheet}`);
    sheetDiv.setAttribute('onclick','viewSheet(event)');
    buttonContainer.append(sheetDiv);
}
function btnHighlighter(no,isYes){
    const btn= document.getElementById(`sheet-${no}`)
    if(isYes){
    btn.style.backgroundColor='#90adcf';
    }
    else{
        btn.style.backgroundColor='white';
    }
   
}
addSheetBTn.addEventListener("click",()=>{
    btnHighlighter(currentSheet)
    //adding sheet btn
    generateNextSheetBtn();
    // saving matrix
    saveMatrix();
    //clean matrix
    createNewMatrix();
    // clean html
    TbodyGen();

})

function saveMatrix(){
    if (localStorage.getItem(ArrMatrix)) {
        // pressing add sheet not for the first time
        let tempArrMatrix = JSON.parse(localStorage.getItem(ArrMatrix));
        tempArrMatrix.push(matrix);
        localStorage.setItem(ArrMatrix, JSON.stringify(tempArrMatrix));
      } else {
        // pressing add sheet for the first time
        let tempArrMatrix = [matrix];
        localStorage.setItem(ArrMatrix, JSON.stringify(tempArrMatrix));
      }
}
function renderMatrix() {
    matrix.forEach((row) => {
      row.forEach((cellObj) => {
        if (cellObj.id) {
          let currentCell = document.getElementById(cellObj.id);
          currentCell.innerText = cellObj.text;
          currentCell.style = cellObj.style;
        }
      });
    });
  }
  
  function viewSheet(event){
    btnHighlighter(prevSheet)
    btnHighlighter(event.target.id.split('-')[1],true)
    // save prev sheet before doing anything
    prevSheet=currentSheet;
    currentSheet=event.target.id.split('-')[1];
    let matrixArr = JSON.parse(localStorage.getItem(ArrMatrix));
    // save my matrix in local storage
    matrixArr[prevSheet-1] = matrix;
    localStorage.setItem(ArrMatrix,JSON.stringify(matrixArr));
  
    // I have updated my virtual memory
    matrix = matrixArr[currentSheet-1];
    // clean my html table
    TbodyGen();
    // render the matrix in html
    renderMatrix();
    prevSheet=event.target.id.split('-')[1];
  }