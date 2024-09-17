var html = $("#template").html();
var template = Handlebars.compile(html);
const excelData =  document.querySelector('form');
const search = document.querySelector('input');

excelData.addEventListener('submit', (e) => {
    e.preventDefault()
    const search_value = search.value;
    console.log("search_value: "+search_value);
    

    messageOne.textContent = 'Loading...'
    messageTwo.textContent = ''

    fetch('/excel_report?address=' +search_value ).then((res) => {
    res.json().then((result) => {
        if(result.error){
            messageOne.textContent = result.error
        }else{
            console.log(result);
            let item = [];
            let res_length = result.length;
            console.log(res_length);
            
            for(let i=0;i<res_length;i++){
                item = "Found  "+result[i].searchValue+ " in file: "+result[i].filePath+ ", sheet: "+result[i].sheetName+ ", cellRef: "+result[i].cellRef;
                // data = result[i];
                console.log(item);
            }
            $('body').append(template({item: item}));
            
        }
        
    })
})
})



var item = [{pagetitle : 'Page 1', text:'content page 1'},
{pagetitle : 'Page 2', text:'content page 2'},
{pagetitle : 'Page 3', text:'content page 3'},
{pagetitle : 'Page 4', text:'content page 4'},
{pagetitle : 'Page 10', text:'content page 10'}];

$('body').append(template({item: item}));