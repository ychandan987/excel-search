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
            console.log(result.length);
            for(let i=0;i<result.length;i++){
                data = result[i];
            }
            messageOne.textContent = data
        }
        
    })
})
})