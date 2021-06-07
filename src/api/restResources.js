


const axios = require("axios");

function getAllData(){
    return axios.get("https://demo.akademie.uni-bremen.de/rest/meta?jsoncallback=acat")
                .then(function (response){
                      return response.data.replace(');',' ').replace('(',' ');
                      

                     

                })
}


getAllData().then(function (response){console.log(JSON.parse(response))})

 