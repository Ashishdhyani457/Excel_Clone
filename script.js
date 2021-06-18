const ps = new PerfectScrollbar("#cells",
 {wheelSpeed: 2,wheelPropagation: true});
for(let i=1;i<=100;i++){
    let str="";
    let n=i;
    while(n>0){
    let rem=n%26;
    if(rem==0){
        str='Z'+str;
        n=Math.floor(n/26)-1;
    }
    else{
    str=String.fromCharCode((rem-1)+65)+str;
    n=Math.floor(n/26);
}
}
$("#columns").append(`<div class="column-name">${str}</div>`);
$("#rows").append(`<div class="row-name">${i}</div>`);

}
let Celldata={
    "Sheet1":{}
};
let selectedSheet="Sheet1";
let totalSheet=1;
let lastlyadddedsheet=1;
let save=true;
let defaultProperties={"font-family" : "Noto Sans",
"font-size":14,
    "text":"",
    "bold":false,
    "italic":false,
    "underlined":false,
    "alignment":"left",
    "bgcolor":"#fff",
    "color":"#444",
    "formula": "",
    "upStream": [],
    "downStream": [],
}
for(let i=1;i<=100;i++){
    let row=$(`<div class="cell-row">`);
    for(let j=1;j<=100;j++){
        row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`);
    }
    $("#cells").append(row);
}
$("#cells").scroll(function(e){
$("#columns").scrollLeft(this.scrollLeft);
$("#rows").scrollTop(this.scrollTop);
});

$(".input-cell").dblclick(function(e){
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    $(this).addClass("selected");
    $(this).attr("contenteditable","true");
    $(this).focus();
})
$(".input-cell").blur(function(e){
    $(this).attr("contenteditable","false");
    updateCelldata("text",$(this).text());
})
function getRowCol(ele){
    let id=$(ele).attr("id");
    let idarray=id.split("-");
    let rowid=parseInt(idarray[1]);
    let colid=parseInt(idarray[3]);
    return [rowid,colid];

}
function gettopleftbottomrightcell(rowid,colid){
    let topcell=$(`#row-${rowid-1}-col-${colid}`);
    let bottomcell=$(`#row-${rowid+1}-col-${colid}`);
    let leftcell=$(`#row-${rowid}-col-${colid-1}`);
    let rightcell=$(`#row-${rowid}-col-${colid+1}`);
    return [topcell,bottomcell,leftcell,rightcell];

}
function unselectcell(ele,e,topcell,bottomcell,leftcell,rightcell){
    if($(ele).attr("contenteditable")=="false"){
        if($(ele).hasClass("top-selected")){
            topcell.removeClass("bottom-selected");
        }
        if($(ele).hasClass("bottom-selected")){
            bottomcell.removeClass("top-selected");
            }
        if($(ele).hasClass("left-selected")){
            leftcell.removeClass("right-selected");
            }
        if($(ele).hasClass("right-selected")){
            rightcell.removeClass("left-selected");
            }
    $(ele).removeClass("selected top-selected bottom-selected left-selected right-selected");
    }
  
}

$(".input-cell").click(function(e){
    let [rowid,colid]=getRowCol(this);
    let[topcell,bottomcell,leftcell,rightcell]=gettopleftbottomrightcell(rowid,colid);
    if($(this).hasClass("selected") && e.ctrlKey){
        unselectcell(this,e,topcell,bottomcell,leftcell,rightcell);

    }else{
    selectCell(this,e,topcell,bottomcell,leftcell,rightcell);}
});


function selectCell(ele,e,topcell,bottomcell,leftcell,rightcell){
if(e.ctrlKey){
    let topselected;
    //top selected or not
if(topcell){
topselected=topcell.hasClass("selected");
}
let bottomselected;
//bottom selected or not
if(bottomcell){
    bottomselected=bottomcell.hasClass("selected");
    }
let leftselected;
//left selected or not
if(leftcell){
    leftselected=leftcell.hasClass("selected");
    }   
let rightselected;
//right selected or not
if(rightcell){
    rightselected=rightcell.hasClass("selected");
    }    

    if(topselected){
        $(ele).addClass("top-selected");
        topcell.addClass("bottom-selected");
    }
    if(bottomselected){
        $(ele).addClass("bottom-selected");
        bottomcell.addClass("top-selected");
    }
    if(leftselected){
        $(ele).addClass("left-selected");
        leftcell.addClass("right-selected");
    }
    if(rightselected){
        $(ele).addClass("right-selected");
        rightcell.addClass("left-selected");
    }

}else{
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
}
$(ele).addClass("selected");
changeHeader(getRowCol(ele));
}
function changeHeader([rowid,colid]){
    let data;
if(Celldata[selectedSheet][rowid-1]&&Celldata[selectedSheet][rowid-1][colid-1]){
    data=Celldata[selectedSheet][rowid-1][colid-1];}
else{
data =defaultProperties;
    }
$(".alignment.selected").removeClass("selected");
$(`.alignment[data-type=${data.alignment}]`).addClass("selected")
addremoveselectfromfontstyle(data,"bold");
addremoveselectfromfontstyle(data,"italic");
addremoveselectfromfontstyle(data,"underlined");
$("#fill-color").css("border-bottom",`4px solid ${data.bgcolor}`);
$("#text-color").css("border-bottom",`4px solid ${data.color}`);
$("#font-family").val(data["font-family"]);
$("#font-size").val(data["font-size"]);
$("#font-family").css("font-family",data["font-family"]);
// console.log(Celldata);
}
function addremoveselectfromfontstyle(data,property){
    if(data[property]){
        $(`#${property}`).addClass("selected");
    }else{
        $(`#${property}`).removeClass("selected");
    }
}

let startcellselected=false;
let startedcell={};
let endcell={};
let scrollxrStarted=false;
let scrollxlStarted=false;
$(".input-cell").mousemove(function(e){
    e.preventDefault();

    if(e.buttons==1){
    //      if(e.pageX>($(window).width()-10) && !scrollxrStarted){
    //             scrollxr();
    //         }
    //         if(e.pageX<(10) && !scrollxlStarted){
    //             scrollxl();
    //         }
        if(!startcellselected){
            let [rowid,colid]=getRowCol(this);
            startedcell={"rowid": rowid, "colid": colid};
            selectAllbetweencells(startedcell,startedcell);
            startcellselected=true;
            $(".input-cell.selected").attr("contenteditable","false");
        }
    
    }else{
        startcellselected=false;
    } 
});

$(".input-cell").mouseenter(function(e){
    if(e.buttons==1){
        // if(e.pageX< ($(window).width()-10) && scrollxrStarted){
        //     clearInterval(scrollXrRinterval);
        //     scrollxrStarted=false;
        // }
        // if(e.pageX>(10) && scrollxlStarted){
        //     clearInterval(scrollxlinterval);
        //     scrollxlStarted=false;
        // }
        let [rowid,colid]=getRowCol(this);
        endcell={"rowid": rowid, "colid": colid};
        selectAllbetweencells(startedcell,endcell);
    }
})
function selectAllbetweencells(start,end){
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected left-selected right-selected");
    for(let i=Math.min(start.rowid, end.rowid);i<=Math.max(start.rowid, end.rowid);i++){
        for(let j=Math.min(start.colid, end.colid);j<=Math.max(start.colid, end.colid);j++){
            let[topcell,bottomcell,leftcell,rightcell]=gettopleftbottomrightcell(i,j);
            selectCell($(`#row-${i}-col-${j}`)[0],{"ctrlKey": true}, topcell, bottomcell, leftcell, rightcell);
        }
    }
}
let scrollXrRinterval;
function scrollxr(){
    scrollxrStarted=true;
    scrollXrRinterval=setInterval(() => {
       $("#cells").scrollLeft($("#cells").scrollLeft()+100);
    }, 100);

}
let scrollxlinterval;
function scrollxl(){
    scrollxlStarted=true;
    scrollxlinterval=setInterval(() => {
       $("#cells").scrollLeft($("#cells").scrollLeft()-100);
    }, 100);

}
$(".data-container").mousemove(function(e){
    e.preventDefault();
    if(e.buttons==1){
        if(e.pageX>($(window).width()-10) && !scrollxrStarted){
               scrollxr();
           }
        if(e.pageX<(10) && !scrollxlStarted){
               scrollxl();
           }
           if(e.pageX< ($(window).width()-10) && scrollxrStarted){
            clearInterval(scrollXrRinterval);
            scrollxrStarted=false;
        }
            if(e.pageX>(10) && scrollxlStarted){
            clearInterval(scrollxlinterval);
            scrollxlStarted=false;
        }
        }
});
$(".data-container").mouseup(function(e){
    clearInterval(scrollXrRinterval);
    clearInterval(scrollxlinterval);
    scrollxrStarted=false;
    scrollxlStarted=false;

});

$(".alignment").click(function(e){
let alignment=$(this).attr("data-type");
$(".alignment.selected").removeClass("selected");
$(this).addClass("selected");
$(".input-cell.selected").css("text-align",alignment);
// $(".input-cell.selected").each(function(index,data){
//    let [rowid,colid]=getRowCol(data);
//         Celldata[rowid - 1][colid - 1].alignment= alignment;
// })
updateCelldata("alignment",alignment);
})
$("#bold").click(function(e){
    setStyle(this,"bold","font-weight","bold")
})
$("#italic").click(function(e){
    setStyle(this,"italic","font-style","italic")
})
$("#underlined").click(function(e){
    setStyle(this,"underlined","text-decoration","underline")
})


function setStyle(ele,property,Key,value){
    if($(ele).hasClass("selected")){
        $(ele).removeClass("selected");
        $(".input-cell.selected").css(Key,"");
    //     $(".input-cell.selected").each(function(index,data){
    //         let [rowid,colid]=getRowCol(data);
    //              Celldata[rowid - 1][colid - 1][property] = false;
    // })
    updateCelldata(property,false);
    }else{
        $(ele).addClass("selected");
        $(".input-cell.selected").css(Key,value);
    //     $(".input-cell.selected").each(function(index,data){
    //         let [rowid,colid]=getRowCol(data);
    //              Celldata[rowid - 1][colid - 1][property] = true;
    // })
    updateCelldata(property,true);
}
}

$(".pick-color").colorPick({
    'initialColor': '#ABCD',
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function() {
      if(this.color!="ABCD"){
        if($(this.element.children()[1]).attr("id") == "fill-color") {
            $(".input-cell.selected").css("background-color",this.color);
            $("#fill-color").css("border-bottom",`4px solid ${this.color}`);
        //     $(".input-cell.selected").each((index,data)=>{
        //         let [rowid,colid]=getRowCol(data);
        //              Celldata[rowid - 1][colid - 1].bgcolor=this.color;
        // })
        updateCelldata("bgcolor",this.color);
        }
        if($(this.element.children()[1]).attr("id") == "text-color") {
            $(".input-cell.selected").css("color",this.color);
            $("#text-color").css("border-bottom",`4px solid ${this.color}`)
        //     $(".input-cell.selected").each((index,data)=>{
        //         let [rowid,colid]=getRowCol(data);
        //              Celldata[rowid - 1][colid - 1].color=this.color;
        // })
        updateCelldata("color",this.color);
        }
    }
    }
  });
$("#fill-color").click(function(e){
    setTimeout(() => {
         $(this).parent().click();
    }, 10);
})
$("#text-color").click(function(e){
    setTimeout(() => {
         $(this).parent().click();
    }, 10);
})
$(".menu-selector").change(function(){
    let value=$(this).val();
    let key=$(this).attr("id");
    if(key=="font-family"){
        $("#font-family").css(key,value)
    }
    if(!isNaN(value)){
        value=parseInt(value);
    }
    $(".input-cell.selected").css(key,value);
    // $(".input-cell.selected").each((index,data)=>{
    //     let [rowid,colid]=getRowCol(data);
    //          Celldata[rowid - 1][colid - 1][key]=value;
//})
updateCelldata(key,value);

})

function updateCelldata(property,value){
    let currcelldata=JSON.stringify(Celldata);
    if(value!=defaultProperties[property]){
        $(".input-cell.selected").each(function(index,data){
         let [rowid,colid]=getRowCol(data);
            if(Celldata[selectedSheet][rowid-1]==undefined){
                Celldata[selectedSheet][rowid-1]={};
                Celldata[selectedSheet][rowid-1][colid-1]={...defaultProperties, "upStream": [], "downStream": []};
                Celldata[selectedSheet][rowid-1][colid-1][property]=value;
            }else{
                if(Celldata[selectedSheet][rowid-1][colid-1]==undefined){
                    Celldata[selectedSheet][rowid-1][colid-1]={};
                    Celldata[selectedSheet][rowid-1][colid-1]={...defaultProperties, "upStream": [], "downStream": []};
                    Celldata[selectedSheet][rowid-1][colid-1][property]=value;             
                }else{
                    Celldata[selectedSheet][rowid-1][colid-1][property]=value;
                }
             
                    }
            });
    }else{
        $(".input-cell.selected").each(function(index,data){
            let [rowid,colid]=getRowCol(data);
        if(Celldata[selectedSheet][rowid-1]&&Celldata[selectedSheet][rowid-1][colid-1]){
            Celldata[selectedSheet][rowid-1][colid-1][property]=value;
            if(JSON.stringify(Celldata[selectedSheet][rowid-1][colid-1])==JSON.stringify(defaultProperties)){
                delete Celldata[selectedSheet][rowid-1][colid-1];
            if(Object.keys(Celldata[selectedSheet][rowid-1]).length==0){
                delete Celldata[selectedSheet][rowid-1];
            }
            }   
        }
    });
    }
    if(save && currcelldata!=JSON.stringify(Celldata)){
        save=false;
    }
  }
  $(".container").click(function(e){
    $(".sheet-options-modal").remove()
  })
  function addSheeetEvents(){
    $(".sheet-tab.selected").on("contextmenu",function(e){
        e.preventDefault();

     
            selectsheet(this);

        $(".sheet-options-modal").remove();
     let modal= $(`<div class="sheet-options-modal">
                 <div class="option sheet-rename">Rename</div>
                 <div class="option sheet-delete">Delete</div>
                 </div>`);
                 modal.css({"left":e.pageX});
 $(".container").append(modal);
 $(".sheet-rename").click(function(){
    let renamemodal=$(`<div class="sheet-modal-parent">
                                   <div class="sheet-rename-modal">
                                       <div class="sheet-modal-title">Rename Sheet</div>
                                           <div class="sheet-modal-input-container">
                                               <span class="sheet-modal-input-title">Rename Sheet To:</span>
                                               <input class="sheet-modal-input" type="text"/>
                                           </div>
                                       <div class="sheet-modal-confirmation">
                                           <div class="button yes-button">OK</div>
                                           <div class="button no-button">Cancel</div>
                                       
                                       </div>
                                   </div>
                               </div>`);
   $(".container").append(renamemodal);
   $(".sheet-modal-input").focus();
   $(".no-button").click(function(e){
       $(".sheet-modal-parent").remove();
   })
   $(".yes-button").click(function(){
        renameSheets();   
   })
   $(".sheet-modal-input").keypress(function(e){
       if(e.key=="Enter"){
           renameSheets();
       }
   })
})

$(".sheet-delete").click(function(e){
    if(totalSheet>1){
let deleteModal=$(`<div class="sheet-modal-parent">
                                <div class="sheet-delete-modal">
                                    <div class="sheet-modal-title">${selectedSheet}</div>
                                        <div class="sheet-modal-detail-container">
                                            <span class="sheet-modal-detail-title">Are you sure?</span>

                                        </div>
                                    <div class="sheet-modal-confirmation">
                                    <div class="button yes-button">
                                        <div class="material-icons delete-icon">delete</div> 
                                        Delete</div>
                                        <div class="button no-button">Cancel</div>
                                    
                                    </div>
                                </div>
                                </div>`);
    $(".container").append(deleteModal);  
    $(".no-button").click(function(e){
        $(".sheet-modal-parent").remove();
    })
    $(".yes-button").click(function(){
        deleteSheets(); 
    })
}else{
    alert("Not possible");
}                         
})


})

 $(".sheet-tab.selected").click(function(){
     
        selectsheet(this);
    
}) 
  }
addSheeetEvents();

$(".add-sheet").click(function(){
    save=false;
lastlyadddedsheet++;
totalSheet++;
Celldata[`Sheet${lastlyadddedsheet}`]={};
// console.log(Celldata);
$(".sheet-tab.selected").removeClass("selected");
$(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet${lastlyadddedsheet}</div>`)
selectsheet();
addSheeetEvents();
$(".sheet-tab.selected")[0].scrollIntoView();
})


function selectsheet(ele){
        if(ele && !$(ele).hasClass("selected")){
            $(".sheet-tab.selected").removeClass("selected");
            $(ele).addClass("selected");
        }
        emptyprevioussheet();
        selectedSheet=$(".sheet-tab.selected").text();
        loadcurrentSheet();
        $("#row-1-col-1").click();
        
}
 function loadcurrentSheet(){
    let data=Celldata[selectedSheet];
    let rowkeys=Object.keys(data);
    for(let i of rowkeys){
        let rowid=parseInt(i);
        let colkeys=Object.keys(data[rowid]);
        for(let j of colkeys){
            let colid=parseInt(j);
            let cell=$(`#row-${rowid+1}-col-${colid+1}`);
            cell.text(data[rowid][colid].text);
            cell.css({ 
                "font-family" : data[rowid][colid]["font-family"],
            "font-size":data[rowid][colid]["font-size"],
                "font-weight":data[rowid][colid].bold ? "bold" : "",
                "font-style":data[rowid][colid].italic ? "italic" : "",
                "text-decoration":data[rowid][colid].underlined ? "underline" : "",
                "text-align":data[rowid][colid].alignment,
                "background-color":data[rowid][colid]["bgcolor"],
                "color":data[rowid][colid].color});
        }
    }
 }
function emptyprevioussheet(){
    let data=Celldata[selectedSheet];
    let rowkeys=Object.keys(data);
    for(let i of rowkeys){
        let rowid=parseInt(i);
        let colkeys=Object.keys(data[rowid]);
        for(let j of colkeys){
            let colid=parseInt(j);
            let cell=$(`#row-${rowid+1}-col-${colid+1}`);
            cell.text("");
            cell.css({"font-family" : "Noto Sans",
            "font-size":14,
                "font-weight":"",
                "font-style":"",
                "text-decoration":"",
                "text-align":"left",
                "background-color":"#fff",
                "color":"#444"})
        }
    }
}
function renameSheets(){
  
let newname=$(".sheet-modal-input").val();
if(newname && !Object.keys(Celldata).includes(newname)){
    let newcelldata={};
    save=false;
    for(let i of Object.keys(Celldata)){
    if(i==selectedSheet){
        newcelldata[newname]=Celldata[selectedSheet];
    }else{
        newcelldata[i]=Celldata[i];
    }
    }
Celldata=newcelldata;


    // Celldata[newname]=Celldata[selectedSheet];
    // delete Celldata[selectedSheet];
    selectedSheet=newname;
    $(".sheet-tab.selected").text(newname);
    $(".sheet-modal-parent").remove();
}else{
    $(".rename-error").remove();
    $(".sheet-modal-input-container").append(`
        <div class="rename-error">Sheet name is not valid or 
        sheet already exists</div>
    `)
}

}

function deleteSheets(){
    $(".sheet-modal-parent").remove();
    let sheetindex=Object.keys(Celldata).indexOf(selectedSheet);
    let currentsheet=$(".sheet-tab.selected");
   if(sheetindex==0){
       selectsheet(currentsheet.next()[0]);
    }else{
    selectsheet(currentsheet.prev()[0]);
    }
    delete Celldata[currentsheet.text()]
       currentsheet.remove();
   totalSheet--;
}

$(".left-scroller,.right-scroller").click(function(){
    let keyarr=Object.keys(Celldata);
    let selectedsheetindex=keyarr.indexOf(selectedSheet);
        if( selectedsheetindex!=0 && $(this).text()=="arrow_left"){
            selectsheet($(".sheet-tab.selected").prev()[0]);
        }else if(selectedsheetindex!=keyarr.length-1 && $(this).text()=="arrow_right"){
            selectsheet($(".sheet-tab.selected").next()[0]);
        }
        $(".sheet-tab.selected")[0].scrollIntoView();
})

$("#menu-file").click(function(){
    let filemodal=$(`<div class="file-modal">
                        <div class="file-options-modal">
                            <div class="close">
                                <div class="material-icons close-icon">arrow_circle_down</div>
                                <div>Close</div>
                            </div>
                            <div class="new">
                                <div class="material-icons new-icon">insert_drive_file</div>
                                <div>New</div>
                            </div>
                            <div class="open">
                                <div class="material-icons open-icon">folder_open</div>
                                <div>Open</div>
                            </div>
                            <div class="save">
                                <div class="material-icons save-icon">save</div>
                                <div>Save</div>
                            </div>
                        </div>
                        <div class="file-recent-modal"></div>
                        <div class="file-transparent"></div>
                    </div>`);
    $(".container").append(filemodal);
    filemodal.animate({
        width : "100vw"
    },300)
    $(".close,.file-transparent,.new,.save,.open").click(function(){
        filemodal.animate({
            width : "0vw"
        },300)
        setTimeout(() => {
            filemodal.remove();
        }, 300);
    });

    $(".new").click(function(){
        if(save){
         newfile();
        }else{
            $(".container").append(`<div class="sheet-modal-parent">
                                <div class="sheet-delete-modal">
                                    <div class="sheet-modal-title">${$(".title").text()}</div>
                                        <div class="sheet-modal-detail-container">
                                            <span class="sheet-modal-detail-title">Do you want to save changes?</span>
                        
                                        </div>
                                    <div class="sheet-modal-confirmation">
                                    <div class="button yes-button">Yes</div>
                                        <div class="button no-button">No</div>
                                    
                                    </div>
                                </div>
                            </div> `)

            $(".yes-button").click(function(e){
                $(".sheet-modal-parent").remove();
            saveFile(true);
            })

            $(".no-button").click(function(e){
                    $(".sheet-modal-parent").remove();
                    newfile();
            })  
         }
    })
    $(".save").click(function(){
        if(!save){
            saveFile();
        }
        
    })
    
    $(".open").click(function(e){
        openFile();
    })
})

function newfile(){
    emptyprevioussheet();
    Celldata={"Sheet1" : {}};
    $(".sheet-tab").remove();
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet1</div>`)
    selectedSheet="Sheet1";
    totalSheet=1;
    lastlyadddedsheet=1;
    $(".title").text("Excel - Clone");
    $("#row-1-col-1").click();
    addSheeetEvents();
}

function saveFile(newClicked){
            $(".container").append(`<div class="sheet-modal-parent">
            <div class="sheet-rename-modal">
                <div class="sheet-modal-title">Save File</div>
                    <div class="sheet-modal-input-container">
                        <span class="sheet-modal-input-title">File Name</span>
                        <input class="sheet-modal-input" value="${$(".title").text()}" type="text" />
                    </div>
                <div class="sheet-modal-confirmation">
                <div class="button yes-button">Save</div>
                    <div class="button no-button">Cancel</div>
                
                </div>
            </div>
        </div>`)

    $(".yes-button").click(function(e){
        $(".title").text($(".sheet-modal-input").val());
        let a = document.createElement("a");
        a.href = `data:application/json,${encodeURIComponent(JSON.stringify(Celldata))}`;
        a.download = $(".title").text() + ".json";
        $(".container").append(a);
        a.click();
        // a.remove();
        save = true;
            })
    $(".no-button,.yes-button").click(function(e){
                $(".sheet-modal-parent").remove();

            if(newClicked){
             newfile();   
            }
                })
        }

        function openFile() {
            let inputFile = $(`<input accept="application/json" type="file" />`);
            $(".container").append(inputFile);
            inputFile.click();
            inputFile.change(function(e) {
                let file = e.target.files[0];
                $(".title").text(file.name.split(".json")[0]);
                let reader = new FileReader();
                reader.readAsText(file);
                reader.onload = () => {
                    emptyprevioussheet();
                    $(".sheet-tab").remove();
                    Celldata = JSON.parse(reader.result);
                    let sheets = Object.keys(Celldata);
                    lastlyadddedsheet = 1;
                    for(let i of sheets) {
                        if(i.includes("Sheet")) {
                            let splittedSheetArray = i.split("Sheet");
                            if(splittedSheetArray.length == 2 && !isNaN(splittedSheetArray[1])) {
                                lastlyadddedsheet = parseInt(splittedSheetArray[1]);
                            }
                        }
                        $(".sheet-tab-container").append(`<div class="sheet-tab selected">${i}</div>`);
                    }
                   addSheeetEvents();
                    $(".sheet-tab").removeClass("selected");
                    $($(".sheet-tab")[0]).addClass("selected");
                    selectedSheet = sheets[0];
                    totalSheet = sheets.length;
                    loadcurrentSheet();
                    inputFile.remove();
                }
            });
        }

        let clipboard = { startCell: [], cellData: {} };
let contentCutted = false;
$("#cut,#copy").click(function (e) {
    if ($(this).text() == "content_cut") {
        contentCutted = true;
    }
    clipboard = { startCell: [], cellData: {} };
    clipboard.startCell = getRowCol($(".input-cell.selected")[0]);
    $(".input-cell.selected").each(function (index, data) {
        let [rowId, colId] = getRowCol(data);
        if (Celldata[selectedSheet][rowId - 1] && Celldata[selectedSheet][rowId - 1][colId - 1]) {
            if (!clipboard.cellData[rowId]) {
                clipboard.cellData[rowId] = {};
            }
            clipboard.cellData[rowId][colId] = { ...Celldata[selectedSheet][rowId - 1][colId - 1] };
        }
    });
  console.log(clipboard.cellData);
});

$("#paste").click(function (e) {
    if (contentCutted) {
        emptyprevioussheet();
    }
    let startCell = getRowCol($(".input-cell.selected")[0]);
    let rows = Object.keys(clipboard.cellData);
    for (let i of rows) {
        let cols = Object.keys(clipboard.cellData[i]);
        for (let j of cols) {
            if (contentCutted) {
                delete Celldata[selectedSheet][i - 1][j - 1];
                if (Object.keys(Celldata[selectedSheet][i - 1]).length == 0) {
                    delete Celldata[selectedSheet][i - 1];
                }
            }

        }
    }
    for (let i of rows) {
        let cols = Object.keys(clipboard.cellData[i]);
        for (let j of cols) {
            let rowDistance = parseInt(i) - parseInt(clipboard.startCell[0]);
            let colDistance = parseInt(j) - parseInt(clipboard.startCell[1]);
            if (!Celldata[selectedSheet][startCell[0] + rowDistance - 1]) {
                Celldata[selectedSheet][startCell[0] + rowDistance - 1] = {};
            }
            Celldata[selectedSheet][startCell[0] + rowDistance - 1][startCell[1] + colDistance - 1] = { ...clipboard.cellData[i][j] };
        }
    }
    loadcurrentSheet();
    if (contentCutted) {
        contentCutted = false;
        clipboard = { startCell: [], cellData: {} };
    }
});
