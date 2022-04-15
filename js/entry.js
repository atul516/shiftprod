/* Dudhichua shift wise entry v2.0
*
* Copyright - GPL Version 3
*
* Library credits: jQuery, jQuery-UI, chosen, popper, jspdf,
*                  html2canvas, html2pdf, jhxlsx, FileSaver 
*
* Author: Atul Pratap Singh (https://github.com/atul516)
*/

var coal_shovel_seam = {};
var ob_shovel_seam = {};
var material_code_coal = 4900000014;
var material_code_ob = 4900000011;
var process_order_purewa_coal = 60008065;
var process_order_turra_coal = 60008066;
var process_order_ob = 70005488;
//JSON in the format as required by jhxlsx library for creating SAP Compatible excel
var dataForSAPCompatibleExcel = [
    {
        "sheetName": "Sheet1",
        "data": []
    }
];
// separate array for special trips
var dataForSpecialTripsExcel = [
    {
        "sheetName": "Sheet1",
        "data": []
    }
];

function check_mandatory_fields_shovel(objj) {
    var flag = false;
    $(objj).find('select.shv').each(function () {
        if (!$(this).val()) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    $(objj).find('input.shv').each(function () {
        if (!$(this).val() || parseFloat($(this).val()) <= 0) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    return flag;
}

function check_mandatory_fields() {
    var flag = false;
    $('#dumper_table select.dumper_n').each(function () {
        if (!$(this).val()) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    $('#dumper_table select.dumper_op').each(function () {
        if (!$(this).val()) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    $('#dumper_table input.dumper_wh').each(function () {
        if (!$(this).val() || parseFloat($(this).val()) <= 0) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    $('#dumper_table select.coal_dump').each(function () {
        var preceding_trips = $(this).parent().parent().find('input.coal_inp');
        var trip_entered = false;
        $(preceding_trips).each(function () {
            if ($(this).val() && parseFloat($(this).val()) > 0) {
                trip_entered = true;
            }
        });
        if (trip_entered && !$(this).val()) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    $('#dumper_table select.ob_dump').each(function () {
        var preceding_trips = $(this).parent().parent().find('input.ob_inp');
        var trip_entered = false;
        $(preceding_trips).each(function () {
            if ($(this).val() && parseFloat($(this).val()) > 0) {
                trip_entered = true;
            }
        });
        if (trip_entered && !$(this).val()) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    return flag;
}

function create_corresponding_dumper_column() {
    if (check_mandatory_fields_shovel($(this).parent().parent().parent())) {
        return;
    }
    let dumper_thead_tr = $('#dumper_table > thead > tr');
    let dumper_tbody_tr = $('#dumper_table > tbody > tr');
    let shovel_table_row = $(this).parent().parent().parent();

    let shovel_unique_id = '';
    let seam_column = $(shovel_table_row).find('select[name="seam[]"]').val();

    if ($(shovel_table_row).find('select[name="material_type[]"]').val() === 'Coal') {
        shovel_unique_id = $(shovel_table_row).find('select[name="shovel_no[]"]').val() + '_' + $(shovel_table_row).find('select[name="shovel_operator[]"]').val() + '_Coal';
        if ($(dumper_thead_tr).find('.' + shovel_unique_id).length > 0) {
            alert('Such shovel combination already exists!');
            $(shovel_table_row).remove();
            return;
        }
        coal_shovel_seam[shovel_unique_id] = [seam_column.split('|')[0], seam_column.split('|')[1]];
        if ($(dumper_thead_tr).find('.coal_dump_column_exists').length <= 0) {
            $("<th><div class='coal_dump_column_exists'>Dump Location<br>(Coal)</div></th>").insertAfter($(dumper_thead_tr).find('.work_hours_dumper_head'));
        }
        $(dumper_thead_tr).find('.coal_dump_column_exists').parent().before("<th class='coal shovel_column " + shovel_unique_id + "'>" + shovel_unique_id.split('_')[0] + "<br>(Coal)</th>");
        let ind_trip = $(dumper_thead_tr).find('.' + shovel_unique_id).index();
        $('#dumper_table > tbody').find('.totalColumn').children().eq(ind_trip - 1).after("<td>0</td>");
        $('#dumper_table > tbody').find('.totalQuantityColumn').children().eq(ind_trip - 1).after("<td>0</td>");
        if ($(dumper_tbody_tr).find('.coal_dump_column_exists').length <= 0) {
            $("<td><select style='width: 110px;' name='coal_dump_location[]' class='searchable coal_dump coal_dump_column_exists'>"
                + "<option value='' selected disabled hidden>Select Dump</option>"
                + "<option value='U022'>East Coal Yard</option>"
                + "<option value='U024'>West Coal Yard</option>"
                + "<option value='U023'>Crusher Yard</option>"
                + "</select></td>").insertAfter(
                    $(dumper_tbody_tr).find('.work_hours_dumper_body')
                );
            let ind_dump = $(dumper_thead_tr).find('.coal_dump_column_exists').parent().index();
            $('#dumper_table > tbody').find('.totalColumn').children().eq(ind_dump - 1).after("<td></td>");
            $('#dumper_table > tbody').find('.totalQuantityColumn').children().eq(ind_dump - 1).after("<td></td>");
        }
        $(dumper_tbody_tr).find('.coal_dump_column_exists').parent().before("<td class='" + shovel_unique_id + "'><input name='" + shovel_unique_id.split('_')[0] + "_Coal_" + shovel_unique_id.split('_')[1] + "[]' class='shovel_dumper_trip coal_inp inp " + "sum_" + shovel_unique_id + "' required='required' maxlength='128' type='number' value='' min='0' data-rule-required='true' data-msg-required='Please enter a valid number'></td>");
    } else if ($(shovel_table_row).find('select[name="material_type[]"]').val() === 'OB') {
        shovel_unique_id = $(shovel_table_row).find('select[name="shovel_no[]"]').val() + '_' + $(shovel_table_row).find('select[name="shovel_operator[]"]').val() + '_OB';
        if ($(dumper_thead_tr).find('.' + shovel_unique_id).length > 0) {
            alert('Such shovel combination already exists!');
            $(shovel_table_row).remove();
            return;
        }
        ob_shovel_seam[shovel_unique_id] = [seam_column.split('|')[0], seam_column.split('|')[1]];
        if ($(dumper_thead_tr).find('.ob_dump_column_exists').length <= 0) {
            $(dumper_thead_tr).find('.dumper_row_delete_button').before("<th><div class='ob_dump_column_exists'>Dump Location<br>(OB)</div></th>");
        }
        $(dumper_thead_tr).find('.ob_dump_column_exists').parent().before("<th class='ob shovel_column " + shovel_unique_id + "'>" + shovel_unique_id.split('_')[0] + "<br>(OB)</th>");
        let ind_trip = $(dumper_thead_tr).find('.' + shovel_unique_id).index();
        $('#dumper_table > tbody').find('.totalColumn').children().eq(ind_trip - 1).after("<td>0</td>");
        $('#dumper_table > tbody').find('.totalQuantityColumn').children().eq(ind_trip - 1).after("<td>0</td>");
        if ($(dumper_tbody_tr).find('.ob_dump_column_exists').length <= 0) {
            $(dumper_tbody_tr).find('.dumper_row_delete_button').before("<td><select style='width: 110px;' name='ob_dump_location[]' class='searchable ob_dump ob_dump_column_exists'>"
                + "<option value='' selected disabled hidden>Select Dump</option>"
                + "<option value='OB01'>OB Dump East</option>"
                + "<option value='OB12'>OB Dump West</option>"
                + "<option value='OB01'>Local OB Dump</option>"
                + "</select></td>");
            let ind_dump = $(dumper_thead_tr).find('.ob_dump_column_exists').parent().index();
            $('#dumper_table > tbody').find('.totalColumn').children().eq(ind_dump - 1).after("<td></td>");
            $('#dumper_table > tbody').find('.totalQuantityColumn').children().eq(ind_dump - 1).after("<td></td>");
        }
        $(dumper_tbody_tr).find('.ob_dump_column_exists').parent().before("<td class='" + shovel_unique_id + "'><input name='" + shovel_unique_id.split('_')[0] + "_OB_" + shovel_unique_id.split('_')[1] + "[]' class='shovel_dumper_trip ob_inp inp " + "sum_" + shovel_unique_id + "' required='required' maxlength='128' type='number' value='' min='0' data-rule-required='true' data-msg-required='Please enter a valid number'></td>");
    }
    //disable for further edit
    $(this).val('Added').removeClass('btn-primary').addClass('btn-success').prop('disabled', true);
    $(this).parent().parent().prevAll().find('input').prop('disabled', true);
    $(this).parent().parent().prevAll().find('select').chosen('destroy').end();
    $(this).parent().parent().prevAll().find('select').prop('disabled', true);

    $('#dumper_table').find(".searchable").chosen().change(setFocusOnNextElement);
    bind_total_event();

    $('td > input').on('keydown', function (e) {
        if (e.which === 13) {
            var element = $(this).parent().next().children('input,select').eq(0);
            if (element.is('input')) {
                window.setTimeout(() => $(element).focus(), 0);
            } else if (element.is('select')) {
                window.setTimeout(() => $(element).trigger('chosen:activate'), 0);
            }
            return false;
        }
    });
}

function bind_total_event() {
    $('#dumper_table .shovel_dumper_trip').bind('keyup change', function () {
        calc_total(this);
    });
}

function calc_total(obj) {
    let $table = $(obj).closest('table');
    let total = 0;
    let total_quantity = 0;
    let ind = $(obj).closest('td').index();
    var name_attr = $(obj).attr("name").split('_');
    let shovel_unique_id = name_attr[0] + '_' + name_attr[2].slice(0, -2) + '_' + name_attr[1];


    if (shovel_unique_id && shovel_unique_id.length > 0) {
        $table.find('tr:not(.totalColumn) .sum_' + shovel_unique_id).each(function () {
            total += +$(this).val();
        });
        $table.find('tr:not(.totalColumn) .sum_' + shovel_unique_id).each(function () {
            var trips = +$(this).val();
            var material_name = $(this).attr("name").split('_')[1];
            var dumper_number = $(this).parent().parent().children('td').eq(0).children('select, input').eq(0).val();
            if (dumper_number !== null) {
                var dumper_factor = get_dumper_factor(dumper_number, material_name);
                total_quantity += parseInt(trips) * parseInt(dumper_factor);
            }
        });
        $table.find('tr.totalColumn').children().eq(ind).html(total);
        $table.find('tr.totalQuantityColumn').children().eq(ind).html(total_quantity);
    }
}

function working_hour_distribution(dumper_working_hours, dumper_shovel_trips) {
    var working_hours = {};
    var total_trips = {};
    for (const key in dumper_shovel_trips) {
        //working_hours_array[key] = {};
        var total_dumper_trips = 0;
        for (const key1 in dumper_shovel_trips[key]) {
            total_dumper_trips += parseInt(dumper_shovel_trips[key][key1]);
        }
        total_trips[key] = total_dumper_trips;
    }
    for (const key in dumper_shovel_trips) {
        working_hours[key] = {};
        for (const key1 in dumper_shovel_trips[key]) {
            working_hours[key][key1] = (parseFloat(dumper_shovel_trips[key][key1]) / parseFloat(total_trips[key])) * parseFloat(dumper_working_hours[key]);
        }
    }
    return working_hours;
}

function get_pdf_report() {
    $('select').each(function () {
        $(this).parent().children('span').remove();
        $(this).after($('<span class="select-print">'
            + $(this).find('option:selected').text() + '</span>'));
    });
    $('input[type="number"],input[type="date"]').each(function () {
        $(this).parent().children('span').remove();
        $(this).after($('<span class="select-print">'
            + $(this).val() + '</span>'));
    });
    $('#print_time').text('Generated on: ' + new Date().toLocaleString() + ' ');
    $("head").append("<link id='printcss' href='css/print.css' type='text/css' rel='stylesheet' />");
    var opt = {
        margin: [0.2, 0.1, 0.2, 0.1],
        filename: $('#date').val() + "_Shift_" + $('#shift').val() + "_" + $('#section').val() + '.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'in', format: 'a4', orientation: 'landscape' }
    };
    html2pdf().set(opt).from(document.body).outputPdf().then(function (pdf) {
        $('#printcss').remove();
        $('.select-print').remove();
    }).save();
}

function populate_data_object_for_sap_excel() {
    var check = check_mandatory_fields();
    if (check == true) {
        alert('ERROR: Empty fields in DUMPER TABLE!');
        return;
    }
    //reinitialize data array to empty
    dataForSAPCompatibleExcel[0].data = [];
    //Create header
    var header = [];
    header.push({ "text": "Plant" });
    header.push({ "text": "Materialcode" });
    header.push({ "text": "Process_Order" });
    header.push({ "text": "Production_Dates" });
    header.push({ "text": "shift" });
    header.push({ "text": "SEAM" });
    header.push({ "text": "Bench" });
    header.push({ "text": "Shovel_number" });
    header.push({ "text": "Operator" });
    header.push({ "text": "Shovel_operating_Time" });
    header.push({ "text": "Dumper_Number" });
    header.push({ "text": "Operator" });
    header.push({ "text": "Dumper_Operating_Time" });
    header.push({ "text": "Dumper_Number_of_Trips" });
    header.push({ "text": "Dumper_Volume" });
    header.push({ "text": "Coal dumped" });
    header.push({ "text": "OB Dumped" });
    header.push({ "text": "Rehandling indicator" });
    header.push({ "text": "DUMPYARD CODE" });
    header.push({ "text": "Breakdown Time" });

    dataForSAPCompatibleExcel[0].data.push(header);

    var dumper_thead_th = $('#dumper_table > thead > tr > th');

    //calculate dumper working hour distribution
    var dumper_working_hours = {};
    //load dumper-wise working hours
    $('#dumper_table input[name="dumper_working_hours[]"]').each(function () {
        if ($(this).val() !== '') {
            dumper_working_hours[$(this).parent().parent().children().first().children('select').eq(0).val() + '_' + $(this).parent().parent().children().eq(1).children('select').eq(0).val()] = $(this).val();
        }
    });

    var dumper_shovel_trips = {};
    //load dumper to shovel-wise trips
    var dumper_tbody_tr = $('#dumper_table > tbody > tr');
    $(dumper_tbody_tr).each(function (index, tr) {
        var unique_dumper = $(tr).children('td').eq(0).children('select').eq(0).val()
            + '_' + $(tr).children('td').eq(1).children('select').eq(0).val();
        dumper_shovel_trips[unique_dumper] = {};
        $(tr).children('td').each(function (index1, td) {
            if ($(td).children('input').eq(0).hasClass('shovel_dumper_trip')
                && $(td).children('select, input').eq(0).val() !== ''
                && $(td).children('select, input').eq(0).val() !== '0'
            ) {
                var unique_shovel = $(td).children('select, input').eq(0).attr("name");
                dumper_shovel_trips[unique_dumper][unique_shovel] = $(td).children('select, input').eq(0).val();
            }
        });
    });

    var dumper_working_hour_distribution = working_hour_distribution(dumper_working_hours, dumper_shovel_trips);

    //calculate shovel working hour distribution
    var shovel_working_hours = {};
    //load dumper-wise working hours
    $('#shovel_table input[name="shovel_working_hours[]"]').each(function () {
        if ($(this).val() !== '') {
            var threeF = $(this).parent().parent().children().first().children('select, input').eq(0).val()
                + '_' + $(this).parent().parent().children().eq(2).children('select, input').eq(0).val()
                + '_' + $(this).parent().parent().children().eq(1).children('select, input').eq(0).val()
                + '[]';

            shovel_working_hours[threeF] = $(this).val();
        }
    });

    var shovel_dumper_trips = {};
    //load dumper to shovel-wise trips
    $('#dumper_table tbody tr').first().find('td').each(function (index, td) {
        if ($(td).children('input').eq(0).hasClass('shovel_dumper_trip')) {
            var unique_shovel = $(td).children('select, input').eq(0).attr("name");
            shovel_dumper_trips[unique_shovel] = {};
            var dumper_columns = $('#dumper_table tr td:nth-child(' + (index + 1) + ')');
            $(dumper_columns).each(function (index1, td1) {
                if ($(td1).children('input').eq(0).hasClass('shovel_dumper_trip')
                    && $(td1).children('select, input').eq(0).val() !== ''
                    && $(td1).children('select, input').eq(0).val() !== '0'
                ) {
                    var unique_dumper = $(td1).parent().children('td').eq(0).children('select').eq(0).val()
                        + '_'
                        + $(td1).parent().children('td').eq(1).children('select').eq(0).val();
                    shovel_dumper_trips[unique_shovel][unique_dumper] = $(td1).children('select, input').eq(0).val();
                }
            });
        }
    });

    var shovel_working_hour_distribution = working_hour_distribution(shovel_working_hours, shovel_dumper_trips);

    $(dumper_thead_th).each(function (index, th) {
        if ($(th).hasClass("shovel_column")) {
            var excelData = [];
            excelData.push({ "text": $('#plant').val() });
            var excelRowToInsert;
            var threeFields;
            var dumper_columns = $('#dumper_table tr td:nth-child(' + (index + 1) + ')');
            $(dumper_columns).each(function (index1, td) {
                if ($(td).children('select, input').eq(0).is('input')
                    && $(td).children('select, input').eq(0).val() !== ''
                    && $(td).children('select, input').eq(0).val() !== '0'
                ) {
                    excelRowToInsert = [];
                    excelRowToInsert = excelData.slice();
                    threeFields = $(td).children('select, input').eq(0).attr("name").split('_');
                    let shovel_unique_id = threeFields[0] + '_' + threeFields[2].slice(0, -2) + '_' + threeFields[1];
                    if (threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": material_code_coal });
                    } else if (threeFields[1] === 'OB') {
                        excelRowToInsert.push({ "text": material_code_ob });
                    }
                    if (threeFields[1] === 'OB') {
                        excelRowToInsert.push({ "text": process_order_ob });
                    } else if (coal_shovel_seam[shovel_unique_id][1].indexOf('TURRA') > -1 && threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": process_order_turra_coal });
                    } else if (coal_shovel_seam[shovel_unique_id][1].indexOf('PURVA') > -1 && threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": process_order_purewa_coal });
                    }
                    let date_to_enter = new Date($('#date').val());
                    let day = date_to_enter.getDate();
                    let month = date_to_enter.getMonth();
                    let year = date_to_enter.getFullYear();
                    month = month + 1;
                    if ((String(day)).length == 1) {
                        day = '0' + day;
                    }
                    if ((String(month)).length == 1) {
                        month = '0' + month;
                    }
                    date_to_enter = day + '.' + month + '.' + year;
                    excelRowToInsert.push({ "text": date_to_enter });
                    excelRowToInsert.push({ "text": $('#shift').val() });
                    if (threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": coal_shovel_seam[shovel_unique_id][0] });
                        excelRowToInsert.push({ "text": coal_shovel_seam[shovel_unique_id][1] });
                        var dump_loc = get_dump_location(
                            'Coal',
                            coal_shovel_seam[shovel_unique_id][0],
                            $('#section').val(),
                            threeFields[0],
                            $(td).parent().children('td').eq(0).children('select, input').eq(0).val()
                        );
                        if (!dump_loc) {
                            dump_loc = $(td).parent().find('select.coal_dump').val();
                        }
                    } else if (threeFields[1] === 'OB') {
                        excelRowToInsert.push({ "text": ob_shovel_seam[shovel_unique_id][0] });
                        excelRowToInsert.push({ "text": ob_shovel_seam[shovel_unique_id][1] });
                        var dump_loc = get_dump_location(
                            'OB',
                            ob_shovel_seam[shovel_unique_id][0],
                            $('#section').val(),
                            threeFields[0],
                            $(td).parent().children('td').eq(0).children('select, input').eq(0).val()
                        );
                        if (!dump_loc) {
                            dump_loc = $(td).parent().find('select.ob_dump').val();
                        }
                    }
                    excelRowToInsert.push({ "text": threeFields[0] });
                    excelRowToInsert.push({ "text": parseInt(threeFields[2]) });
                    //shovel operating time
                    if (threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": parseInt(shovel_working_hour_distribution[$(td).children('select, input').eq(0).attr("name")][$(td).parent().children('td').eq(0).children('select, input').eq(0).val() + '_' + $(td).parent().children('td').eq(1).children('select, input').eq(0).val()] * 60) });
                    } else if (threeFields[1] === 'OB') {
                        excelRowToInsert.push({ "text": parseInt(shovel_working_hour_distribution[$(td).children('select, input').eq(0).attr("name")][$(td).parent().children('td').eq(0).children('select, input').eq(0).val() + '_' + $(td).parent().children('td').eq(1).children('select, input').eq(0).val()] * 60) });
                    }
                    excelRowToInsert.push({ "text": $(td).parent().children('td').eq(0).children('select, input').eq(0).val() });
                    excelRowToInsert.push({ "text": parseInt($(td).parent().children('td').eq(1).children('select, input').eq(0).val()) });
                    excelRowToInsert.push({ "text": parseInt(dumper_working_hour_distribution[$(td).parent().children('td').eq(0).children('select').eq(0).val() + '_' + $(td).parent().children('td').eq(1).children('select').eq(0).val()][$(td).children('select, input').eq(0).attr("name")] * 60) });
                    var trips = $(td).children('select, input').eq(0).val();
                    excelRowToInsert.push({ "text": parseInt(trips) });
                    var dumper_factor = get_dumper_factor($(td).parent().children('td').eq(0).children('select, input').eq(0).val(), threeFields[1]);
                    excelRowToInsert.push({ "text": parseInt(dumper_factor) });
                    var dumper_tonnage = trips * dumper_factor;
                    if (threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": parseInt(dumper_tonnage) });
                        excelRowToInsert.push({ "text": '' });
                    } else if (threeFields[1] === 'OB') {
                        excelRowToInsert.push({ "text": '' });
                        excelRowToInsert.push({ "text": parseInt(dumper_tonnage) });
                    }
                    excelRowToInsert.push({ "text": '' });
                    excelRowToInsert.push({ "text": dump_loc });
                    excelRowToInsert.push({ "text": '' });
                    dataForSAPCompatibleExcel[0].data.push(excelRowToInsert);
                }
            });
        }
    });
    populate_special_trips_table();
    $('#dumperwise_entry').hide('slide', { direction: 'left' }, 1000);
    $('#special_trips_div').show('slide', { direction: 'right' }, 1000);
}

function populate_data_object_for_special_trips_excel() {
    //reinitialize data array to empty
    dataForSpecialTripsExcel[0].data = [];
    //Create header
    let header = [];
    header.push({ "text": "Oprtr/Vend" });
    header.push({ "text": "Name" });
    header.push({ "text": "Trips" });
    header.push({ "text": "First hr" });
    header.push({ "text": "Mid hr" });
    header.push({ "text": "Last hr" });

    dataForSpecialTripsExcel[0].data.push(header);

    $('#special_trips_table > tbody > tr').each(function(index, tr) {
        let Oprtr = $(tr).find('td span').eq(0).text();
        let Name = $(tr).find('td span').eq(1).text();
        let Trips = $(tr).find('td span').eq(2).text();
        let first_hr = $(tr).find('td > input').eq(0).val();
        let mid_hr = $(tr).find('td > input').eq(1).val();
        let last_hr = $(tr).find('td > input').eq(2).val();
        //data to insert
        let data = [];
        data.push({ "text": Oprtr });
        data.push({ "text": Name });
        data.push({ "text": Trips });
        data.push({ "text": first_hr });
        data.push({ "text": mid_hr });
        data.push({ "text": last_hr });

        dataForSpecialTripsExcel[0].data.push(data);
    });
}

function get_dump_location(material_type, seam, section, shovel_name, dumper_name) {
    /*
    --> All OB: OB12 fix
    --> Coal turra (Km and CN) : U023 fix
    --> Coal pureva bottom : east shovels (12, 15, Laxman): ANY DUMPER: U022 (east yard) fix
    --> Coal pureva bottom : west shovels ( 17/19/16/Laxman): Default me U023 (CHP) for 100te dumpers + U024 (west yard) for 190 te (CAT)
    Manual correction needed only for west section.. if mentioned in the report of West section (each shift), then manually change..
    */
    if (material_type === 'OB') {
        return 'OB12';
    } else if (material_type === 'Coal'
        && seam.indexOf('TURRA') > -1
        && (dumper_name.indexOf('KM') > -1 || dumper_name.indexOf('CN') > -1)
    ) {
        return 'U023';
    } else if (material_type === 'Coal'
        && seam.indexOf('PURVA-BOTM-COAL') > -1
        && section === 'East'
    ) {
        return 'U022';
    } else if (material_type === 'Coal'
        && seam.indexOf('PURVA-BOTM-COAL') > -1
        && section === 'West'
        && (dumper_name.indexOf('KM') > -1 || dumper_name.indexOf('CN') > -1)
    ) {
        return 'U023';
    } else if (material_type === 'Coal'
        && seam.indexOf('PURVA-BOTM-COAL') > -1
        && section === 'West'
        && dumper_name.indexOf('CAT') > -1
    ) {
        return 'U024';
    } else {
        return null;
    }
}

function go_back() {
    $('#special_trips_div').hide('slide', { direction: 'right' }, 1000);
    $('#dumperwise_entry').show('slide', { direction: 'left' }, 1000);
}

function go_back_1() {
    $('#excel_buttons').hide('slide', { direction: 'right' }, 1000);
    $('#special_trips_div').show('slide', { direction: 'left' }, 1000);
}

function go_forward() {
    populate_data_object_for_special_trips_excel();
    $('#special_trips_div').hide('slide', { direction: 'left' }, 1000);
    $('#excel_buttons').show('slide', { direction: 'right' }, 1000);
}

function get_sap_compatible_excel() {
    if (dataForSAPCompatibleExcel[0].data.length <= 0) {
        alert('Error: Data object is not populated yet.');
        return;
    }
    var options = {
        fileName: $('#date').val() + "_Shift_" + $('#shift').val() + "_" + $('#section').val()
    };
    Jhxlsx.export(dataForSAPCompatibleExcel, options);
}

function get_special_trips_excel() {
    if (dataForSpecialTripsExcel[0].data.length <= 0) {
        alert('Error: Data object is not populated yet.');
        return;
    }
    var options = {
        fileName: "Special_Trips_" + $('#date').val() + "_Shift_" + $('#shift').val() + "_" + $('#section').val()
    };
    Jhxlsx.export(dataForSpecialTripsExcel, options);
}

function populate_special_trips_table() {
    var unique_operator_for_special_trips = {};
    //Get operator names from shovel table
    $('#shovel_table > tbody > tr').each(function (index, tr) {
        var operator = $(tr).children('td').eq(1).children('select').eq(0);
        if ($(operator).length > 0 && $(operator).val() !== '') {
            if (!($(operator).val() in unique_operator_for_special_trips)) {
                unique_operator_for_special_trips[$(operator).val()] = {};
                unique_operator_for_special_trips[$(operator).val()]['Oprtr/Vend'] = $(operator).val();
                unique_operator_for_special_trips[$(operator).val()]['Name'] = $(operator).find("option:selected").text().split('--')[0].trim();
                var total_trips = 0;
                unique_operator_for_special_trips[$(operator).val()]['Trips'] = total_trips;
            }
        }
    });
    //get operator names from dumper table as well as count their total trips
    $('#dumper_table > tbody > tr').each(function (index, tr) {
        var operator = $(tr).children('td').eq(1).children('select').eq(0);
        if ($(operator).length > 0 && $(operator).val() !== '') {
            if (!($(operator).val() in unique_operator_for_special_trips)) {
                unique_operator_for_special_trips[$(operator).val()] = {};
                unique_operator_for_special_trips[$(operator).val()]['Oprtr/Vend'] = $(operator).val();
                unique_operator_for_special_trips[$(operator).val()]['Name'] = $(operator).find("option:selected").text().split('--')[0].trim();
                var total_trips = 0;
                $(tr).find('.shovel_dumper_trip').each(function () {
                    if ($(this).val() !== '') {
                        total_trips += parseInt($(this).val());
                    }
                });
                unique_operator_for_special_trips[$(operator).val()]['Trips'] = total_trips;
            } else if ($(operator).val() in unique_operator_for_special_trips) {
                var total_trips = 0;
                $(tr).find('.shovel_dumper_trip').each(function () {
                    if ($(this).val() !== '') {
                        total_trips += parseInt($(this).val());
                    }
                });
                unique_operator_for_special_trips[$(operator).val()]['Trips'] += total_trips;
            }
        }
    });

    $('#special_trips_table > tbody').children().remove();
    $.each(unique_operator_for_special_trips, function (key, value) {
        let operator_no = unique_operator_for_special_trips[key]['Oprtr/Vend'];
        let operator_name = unique_operator_for_special_trips[key]['Name'];
        let total_trips = unique_operator_for_special_trips[key]['Trips'];
        const tr = `
    <tr>
      <td>
        <span>${operator_no}</span>
      </td>
      <td>
        <span>${operator_name}</span>
      </td>
      <td>
        <span>${total_trips}</span>
      </td>
      <td>
        <input name="first_hr[]" class='inp1 shv' required="required"
            maxlength="128" type="number" value="" placeholder=""
            data-rule-required="true" data-msg-required="Please enter a valid number">
      </td>
      <td>
        <input name="mid_hr[]" class='inp1 shv' required="required"
            maxlength="128" type="number" value="" placeholder=""
            data-rule-required="true" data-msg-required="Please enter a valid number">
      </td>
      <td>
        <input name="last_hr[]" class='inp1 shv' required="required"
            maxlength="128" type="number" value="" placeholder=""
            data-rule-required="true" data-msg-required="Please enter a valid number">
      </td>
    </tr>
    `;
        $('#special_trips_table > tbody').append(tr);
    });
}

function get_dumper_factor(dumper_number, material_type) {
    var df;
    /*
    CN-01 TO CN-36 FOR COAL=45 FOR OB 32
    C-SERIES & K/D SERIES COAL 40 & OB 27
    TX SERIES COAL 55 OB 37
    KM SERIES FOR COAL=45 FOR OB 32
    CAT SERIES COAL 75 OB 50
    */
    if (dumper_number.indexOf('CN-') > -1) {
        df = material_type == 'Coal' ? 45 : 32;
    } else if (dumper_number.indexOf('C-') > -1
        || dumper_number.indexOf('K-') > -1
        || dumper_number.indexOf('D-') > -1) {
        df = material_type == 'Coal' ? 40 : 27;
    } else if (dumper_number.indexOf('TX-') > -1) {
        df = material_type == 'Coal' ? 55 : 37;
    } else if (dumper_number.indexOf('KM-') > -1) {
        df = material_type == 'Coal' ? 45 : 32;
    } else if (dumper_number.indexOf('CAT-') > -1) {
        df = material_type == 'Coal' ? 75 : 50;
    }
    return df;
}

function setFocusOnNextElement() {
    var element = $(this).parent().next().children('input,select').eq(0);
    if (element.is('input')) {
        window.setTimeout(() => $(element).focus(), 0);
    } else if (element.is('select')) {
        window.setTimeout(() => $(element).trigger('chosen:activate'), 0);
    }
    return false;
}

function updateSeam() {
    var options = '';
    if ($(this).val() == 'Coal') {
        options = "<option value='' selected disabled hidden>Select Seam</option> \
        <option value='PURVA-TOP-COAL|PURVA-TOP-COAL EAST'>Pureva Top Coal-East</option> \
        <option value='PURVA-TOP-COAL|PURVA-TOP-COAL WEST'>Pureva Top Coal-West</option> \
        <option value='PURVA-BOTM-COAL|PURVA-BOTM-COAL EAST'>Pureva Bottom Coal-East</option> \
        <option value='PURVA-BOTM-COAL|PURVA-BOTM-COAL WEST'>Pureva Bottom Coal-West</option> \
        <option value='TURRA COAL|TURRA COAL EAST'>Turra Coal-East</option> \
        <option value='TURRA COAL|TURRA COAL WEST'>Turra Coal-West</option>";
    } else if ($(this).val() == 'OB') {
        options = "<option value='' selected disabled hidden>Select Seam</option> \
        <option value='TURRA BAND-OB|TURRA BAND-OB WEST'>OB-Turra Band West</option> \
        <option value='TURRA BAND-OB|TURRA BAND-OB EAST'>OB-Turra Band East</option> \
        <option value='SM BAND-OB|SM BAND-OB WEST'>OB-SM Band West</option> \
        <option value='SM BAND-OB|SM BAND-OB EAST'>OB-SM Band East</option> \
        <option value='B/W PURTOP & BOT-OB|B/W PURTOP& BOT-OB-W'>OB-Above Purewa Bottom-West</option> \
        <option value='B/W PURTOP & BOT-OB|B/W PURTOP& BOT-OB-E'>OB-Above Purewa Bottom-East</option>"
    }
    $(this).parent().next().children('select').chosen('destroy').end();
    $(this).parent().next().children('select').eq(0).empty().append(options);
    $(this).parent().next().children('select').eq(0).chosen().change(setFocusOnNextElement);
}

function delete_synchrnous_row_and_column() {
    var tbody = $(this).closest("tbody");
    if ($(tbody).children("tr").length > 1) {
        let dumper_thead_tr = $('#dumper_table > thead > tr');
        let dumper_tbody = $('#dumper_table > tbody');
        let shovel_table_row = $(this).closest('tr');
        let shovel_unique_id = $(shovel_table_row).find('select[name="shovel_no[]"]').val() + '_' + $(shovel_table_row).find('select[name="shovel_operator[]"]').val() + '_' + $(shovel_table_row).find('select[name="material_type[]"]').val();
        shovel_unique_id = shovel_unique_id.trim();
        let ind_trip = $(dumper_thead_tr).find('.' + shovel_unique_id).index();
        $(shovel_table_row).remove();
        if ($(dumper_thead_tr).find('.' + shovel_unique_id).length <= 0) {
            return;
        }
        $('#dumper_table > tbody').find('.totalColumn').children().eq(ind_trip).remove();
        $('#dumper_table > tbody').find('.totalQuantityColumn').children().eq(ind_trip).remove();
        $(dumper_tbody).find('.' + shovel_unique_id).remove();
        $(dumper_thead_tr).find('.' + shovel_unique_id).remove();
        if ($(dumper_thead_tr).find('.coal').length <= 0 && $(dumper_thead_tr).find('.coal_dump_column_exists').length > 0) {
            let ind_dump = $(dumper_thead_tr).find('.coal_dump_column_exists').parent().index();
            $('#dumper_table > tbody').find('.totalColumn').children().eq(ind_dump).remove();
            $('#dumper_table > tbody').find('.totalQuantityColumn').children().eq(ind_dump).remove();
            $(dumper_thead_tr).find('.coal_dump_column_exists').parent().remove();
            $(dumper_tbody).find('.coal_dump_column_exists').parent().remove();
        }
        if ($(dumper_thead_tr).find('.ob').length <= 0 && $(dumper_thead_tr).find('.ob_dump_column_exists').length > 0) {
            let ind_dump = $(dumper_thead_tr).find('.ob_dump_column_exists').parent().index();
            $('#dumper_table > tbody').find('.totalColumn').children().eq(ind_dump).remove();
            $('#dumper_table > tbody').find('.totalQuantityColumn').children().eq(ind_dump).remove();
            $(dumper_thead_tr).find('.ob_dump_column_exists').parent().remove();
            $(dumper_tbody).find('.ob_dump_column_exists').parent().remove();
        }
    }
}
function add_row_to_shovel_table() {
    let temp = $('#shovel_row').clone();
    $('#shovel_table > tbody').append($(temp).html());
    $('#shovel_table > tbody')
        .children()
        .last()
        .find('.searchable').chosen().change(setFocusOnNextElement);
    $('#shovel_table > tbody')
        .children()
        .last()
        .find('select[name="material_type[]"]').on('change', updateSeam);
    $('#shovel_table > tbody')
        .children()
        .last()
        .find(".create-dumper-column").on('click', create_corresponding_dumper_column);
    $('#shovel_table > tbody')
        .children()
        .last()
        .find(".delete_row1").on('click', delete_synchrnous_row_and_column);
}
function add_row_to_dumper_table() {
    let tbody = $('#dumper_table > tbody');
    let dumper_thead_tr = $('#dumper_table > thead > tr');
    if ($(dumper_thead_tr).children().length <= 4) {
        var temp = $('#dumper_row').clone();
        temp = $(temp).html();
    } else {
        $(tbody).children().first().find('.searchable').chosen('destroy').end();
        var temp = $(tbody).children().first().clone();
        $(tbody).children().first().find('.searchable').chosen().change(setFocusOnNextElement);
        $(temp).find('input').not(':input[type=button]').val('');
    }
    $(tbody)
        .children()
        .eq(-2)
        .before(temp);
    $(tbody)
        .children()
        .eq(-3)
        .find('.searchable').chosen().change(setFocusOnNextElement);
    $(tbody)
        .children()
        .eq(-3)
        .find(".delete_row2").on('click', delete_dummper_row);

    bind_total_event();
    $('td > input').on('keydown', function (e) {
        if (e.which === 13) {
            var element = $(this).parent().next().children('input,select').eq(0);
            if (element.is('input')) {
                window.setTimeout(() => $(element).focus(), 0);
            } else if (element.is('select')) {
                window.setTimeout(() => $(element).trigger('chosen:activate'), 0);
            }
            return false;
        }
    });
}
function delete_dummper_row() {
    var tbody = $(this).closest("tbody");
    if ($(tbody).children("tr").length > 3) {
        $(this).closest('tr').remove();
        $('#dumper_table > tbody > tr').first().find('.shovel_dumper_trip').each(function () {
            calc_total(this);
        });
    }
}
$(document).ready(function () {
    var now = new Date();
    var day = ("0" + now.getDate()).slice(-2);
    var month = ("0" + (now.getMonth() + 1)).slice(-2);

    var today = now.getFullYear() + "-" + (month) + "-" + (day);
    $('#date').val(today);

    add_row_to_shovel_table();
    let temp = $('#dumper_total_rows').clone();
    $('#dumper_table > tbody').append($(temp).html());
    add_row_to_dumper_table();

    $(".add_row1").on('click', add_row_to_shovel_table);
    $(".add_row2").on('click', add_row_to_dumper_table);
    $("#validate2").on('click', check_mandatory_fields);
    $("#populate_data_object").on('click', populate_data_object_for_sap_excel);
    $("#get_sap_compatible_excel").on('click', get_sap_compatible_excel);
    $("#get_special_trips_excel").on('click', get_special_trips_excel);
    $("#go_back").on('click', go_back);
    $("#go_back_1").on('click', go_back_1);
    $("#go_forward").on('click', go_forward);
    $("#get_pdf_report").on('click', get_pdf_report);
});
