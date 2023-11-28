window.addEventListener('DOMContentLoaded', (event) =>{
    getVisitCount();
})

const functionApiUrl = 'https://getresumecountercmc.azurewebsites.net/api/GetResumeCounter?code=q0WYl_bwbmjjax2bzOSLONIEbIfkyviN2Gnm-tdQHZAkAzFu5nZ-Yg==';
const localfunctionApi = 'http://localhost:7071/api/GetResumeCounter';
//const functionApi = 'http://localhost:7071/api/GetResumeCounter';
const getVisitCount = () => {
    let count = 30;
    fetch(functionApiUrl).then(response => {
        return response.json()
    }).then(response => {
        console.log("Website called function API.");
        count = response.count;
        document.getElementById("counter").innerText = count;
    }).catch(function(error){
        console.log(error);
    });
    return count;
}