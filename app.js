/**
 * Created by korhanozbek on 3.12.2016.
 */
let cheerio = require('cheerio');
var request = require("request");
var excelbuilder = require('msexcel-builder');

var categoriesList = [];
var companiesList = [];


let websiteURL = 'http://www.**.com';

var finalCallback = function () {
    console.log("this is the final.");
    saveExcel(companiesList);
};

getCategoriesListDatas(function () {
    getCompaniesListDatas(finalCallback);
});


function getCategoriesListDatas(callback) {


    request({uri: websiteURL
    }, function(error, response, body) {

        let main = cheerio.load(body);

        main("#tamamiAciklama .col-xs-6 a").each(function (id, element) {

            const catTitle = element.attribs.title;

            const catLink = websiteURL + "/" + element.attribs.href;

            categoriesList[id] = {title : catTitle, link : catLink};


        });

        callback();

    });
}

function getCompaniesListDatas(finalCallback) {

    requestCompaniesList(0, categoriesList.length, function() {
        getCompanyInfoDatas(0, companiesList.length, finalCallback)
    });

}


function requestCompaniesList(i, length, callback) {

    if(i<length) {

        console.log("i = " + i);

        var cat = categoriesList[i];

        request({uri: cat.link
        }, function(error, response, body) {

            let category = cheerio.load(body);

            category(".car-title a").each(function (id, element) {
                var _id = companiesList.length;
                companiesList[_id] = {categoryTitle: cat.title, title : element.attribs.title, link : websiteURL + "/" + element.attribs.href};
            });

            requestCompaniesList(i+1, categoriesList.length, callback);
        });

    }else {
        callback();
    }
}


function getCompanyInfoDatas(j, length, finalCallback) {

    if(j<length) {

        console.log("j = " + j);

        var comp = companiesList[j];

        request({uri: comp.link
        }, function(error, response, body) {

            let company = cheerio.load(body);

            company(".col-md-12 .fdtxt").each(function (id, element) {

                var phone = "";
                try{
                    phone = element.children[0].data;
                }catch(e) {
                    phone = ""
                }

                companiesList[j] = {categoryTitle: comp.categoryTitle, title : comp.title, link : comp.link,
                    phone: phone
                };
            });

            getCompanyInfoDatas(j+1, companiesList.length, finalCallback);

        });

    }else {
        finalCallback();
    }
}

function saveExcel(companyData) {
    // Create a new workbook file in current working-path
    var workbook = excelbuilder.createWorkbook('./', 'malatyaFirmaRehberi.xlsx');

    // Create a new worksheet with 10 columns and 12 rows
    var sheet1 = workbook.createSheet('malatyaFirmaRehberi', 7, companyData.length+10);

    // Fill some data
    sheet1.set(2, 1, 'Kategorisi');
    sheet1.set(3, 1, 'Adı');
    sheet1.set(4, 1, 'Telefonu');

    sheet1.width(2, 30);
    sheet1.width(3, 50);
    sheet1.width(4, 20);

    sheet1.width(6, 50);
    sheet1.set(6, 1, 'Toplam Bulunan Firma Sayısı = ' + companyData.length);

    var baslangicSatiri = 2;
    //önce sutun sonra satir yaziyor.
    for (var i = 0; i < companyData.length; i++) {

        var data = companyData[i];
        var satir = i + baslangicSatiri;

        sheet1.set(2, satir, data.categoryTitle);
        sheet1.set(3, satir, data.title.valueOf());
        sheet1.set(4, satir, data.phone);

    }
    // Save it
    workbook.save(function(ok){
        if (!ok)
            workbook.cancel();
        else
            console.log('congratulations, your workbook created');
    });
}

