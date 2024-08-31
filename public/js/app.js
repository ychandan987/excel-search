const excelData =  document.querySelector('form');
const search = document.querySelector('input');
const messageOne = document.querySelector('#message-1');
const messageTwo = document.querySelector('#message-2');




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
            let data1 = [];
            let res_length = result.length;
            console.log(res_length);
            
            for(let i=0;i<res_length;i++){
                // data = "Found  "+result[i].searchValue+ " in file: "+result[i].filePath+ ", sheet: "+result[i].sheetName+ ", cellRef: "+result[i].cellRef;
                data1 = "<table>"
                "<td>"+result[i].searchValue+ "</td>"
                "<td>"+result[i].filePath+ "</td>"
                "<td>"+result[i].sheetName+ "</td>"
                "<td>"+result[i].cellRef+ "</td>"
           "</table>"
                // data = result[i];
                console.log(data1);
            }
            messageOne.textContent = data1
            
        }
        
    })
})
})