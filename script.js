 // dimension of  excel sheet-------- 
 let column=26;
 let Row=100;
 let currentCell;

 //assecing element---------------
const TheadRow =document.getElementById("TheadRow");
const Tbody = document.getElementById("tbody");
const currentCellHeading=document.querySelector(".currentCell")

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
        cell.addEventListener("focus",(event)=>focusCurrentElement(event.target))
    }
    tableRow.append(cell);
   }
}
colGen("th",TheadRow,true)

// getting current cell and setting color of respective header
 function focusCurrentElement(cell){
    currentCell=cell;
    currentCellHeading.innerText=`${cell.id}`
    setHeaderColor(cell.id[0],cell.id.substring(1));
 }

 // function for setting color of header
 function setHeaderColor(colID,rowID)
 {

    const colH = document.getElementById(colID);
    const rowH = document.getElementById(rowID);
    colH.style.backgroundColor ="green";
    rowH.style.backgroundColor ="green";

 }

 // genreating row.........
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