/* Dudhichua shift wise entry
*
* Generic Copyright
*
* Libraries used: jQuery, jspdf, html2canvas, html2pdf, jhxlsx
* Author: Atul Pratap Singh (https://github.com/atuldch)
*/

var date = 'today';
var shift = 1;
var section = 1;
var coal_shovel_seam = {};
var ob_shovel_seam = {};
var material_code_coal = 4900000014;
var material_code_ob = 4900000011;
var process_order_purewa_coal = 60006912;
var process_order_turra_coal = 60006914;
var process_order_ob = 70004901;

function check_mandatory_fields_shovel() {
    var flag = false;
    $('#shovel_table > tbody > tr select.shv').each(function () {
        if (!$(this).val()) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    $('#shovel_table > tbody > tr input.shv').each(function () {
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
        $(preceding_trips).each(function() {
            if ($(this).val() && parseFloat($(this).val()) > 0) {
                trip_entered = true;
            }
        });
        if(trip_entered && !$(this).val()) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    $('#dumper_table select.ob_dump').each(function () {
        var preceding_trips = $(this).parent().parent().find('input.ob_inp');
        var trip_entered = false;
        $(preceding_trips).each(function() {
            if ($(this).val() && parseFloat($(this).val()) > 0) {
                trip_entered = true;
            }
        });
        if(trip_entered && !$(this).val()) {
            flag = true;
            $(this).parent().addClass("error_in_field");
        } else {
            $(this).parent().removeClass("error_in_field");
        }
    });
    return flag;
}

//JSON for data
var dataForPage = [
    {
        "sheetName": "Sheet1",
        "data": []
    }
];

function create_corresponding_dumper_column() {
    if (check_mandatory_fields_shovel()) {
        return;
    }
    let dumper_thead_tr = $('#dumper_table > thead > tr');
    let dumper_tbody_tr = $('#dumper_table > tbody > tr');
    let shovel_table_row = $(this).parent().parent().parent();

    $('#dumperwise_entry').fadeOut(100);

    let shovel_unique_id = '';
    let seam_column = $(shovel_table_row).find('select[name="seam[]"]').val();

    if ($(shovel_table_row).find('select[name="material_type[]"]').val() === 'Coal' ) {
        shovel_unique_id = $(shovel_table_row).find('select[name="shovel_no[]"]').val() + '_' + $(shovel_table_row).find('select[name="shovel_operator[]"]').val() + '_Coal';       
        if ($(dumper_thead_tr).find('.' + shovel_unique_id).length > 0) {
            $('#dumperwise_entry').fadeIn(300);
            return;
        }
        coal_shovel_seam[shovel_unique_id] = [seam_column.split('|')[0], seam_column.split('|')[1]];
        if ($(dumper_thead_tr).find('.coal_dump_column_exists').length <= 0) {
            $("<th><div class='coal_dump_column_exists'>Dump Location<br>(Coal)</div></th>").insertAfter($(dumper_thead_tr).find('.work_hours_dumper_head'));
        }
        $(dumper_thead_tr).find('.coal_dump_column_exists').parent().before("<th class='shovel_column " + shovel_unique_id + "'>" + shovel_unique_id.split('_')[0] + "<br>(Coal)</th>");
        if ($(dumper_tbody_tr).find('.coal_dump_column_exists').length <= 0) {
            $("<td><select style='width: 110px;' name='coal_dump_location[]' class='searchable coal_dump coal_dump_column_exists'>"
                + "<option value='' selected disabled hidden>Select Dump</option>"
                + "<option value='U022'>East Coal Yard</option>"
                + "<option value='U024'>West Coal Yard</option>"
                + "<option value='U023'>Crusher Yard</option>"
                + "</select></td>").insertAfter(
                    $(dumper_tbody_tr).find('.work_hours_dumper_body')
                );
        }
        $(dumper_tbody_tr).find('.coal_dump_column_exists').parent().before("<td class='" + shovel_unique_id + "'><input name='" + shovel_unique_id.split('_')[0] + "_Coal_" + shovel_unique_id.split('_')[1] + "[]' class='shovel_dumper_trip coal_inp inp " + "sum" + "4" + "' required='required' maxlength='128' type='number' value='' min='0' data-rule-required='true' data-msg-required='Please enter a valid number'></td>");
    } else if ($(shovel_table_row).find('select[name="material_type[]"]').val() === 'OB') {
        shovel_unique_id = $(shovel_table_row).find('select[name="shovel_no[]"]').val() + '_' + $(shovel_table_row).find('select[name="shovel_operator[]"]').val() + '_OB';
        if ($(dumper_thead_tr).find('.' + shovel_unique_id).length > 0) {
            $('#dumperwise_entry').fadeIn(300);
            return;
        }
        ob_shovel_seam[shovel_unique_id] = [seam_column.split('|')[0], seam_column.split('|')[1]];
        if ($(dumper_thead_tr).find('.ob_dump_column_exists').length <= 0) {
            $(dumper_thead_tr).append("<th><div class='ob_dump_column_exists'>Dump Location<br>(OB)</div></th>");
        }
        $(dumper_thead_tr).find('.ob_dump_column_exists').parent().before("<th class='shovel_column " + shovel_unique_id + "'>" + shovel_unique_id.split('_')[0] + "<br>(OB)</th>");
        if ($(dumper_tbody_tr).find('.ob_dump_column_exists').length <= 0) {
            $(dumper_tbody_tr).find('.dumper_row_delete_button').before("<td><select style='width: 110px;' name='ob_dump_location[]' class='searchable ob_dump ob_dump_column_exists'>"
                + "<option value='' selected disabled hidden>Select Dump</option>"
                + "<option value='OB01'>OB Dump East</option>"
                + "<option value='OB12'>OB Dump West</option>"
                + "<option value='OB01'>Local OB Dump</option>"
                + "</select></td>");
        }
        $(dumper_tbody_tr).find('.ob_dump_column_exists').parent().before("<td class='" + shovel_unique_id + "'><input name='" + shovel_unique_id.split('_')[0] + "_OB_" + shovel_unique_id.split('_')[1] + "[]' class='shovel_dumper_trip ob_inp inp " + "sum" + "4" + "' required='required' maxlength='128' type='number' value='' min='0' data-rule-required='true' data-msg-required='Please enter a valid number'></td>");
    }
    $(this).val('Added').removeClass('btn-primary').addClass('btn-success');
    $('#dumperwise_entry').fadeIn(300);

/*
    var totals_html = '<tr class=totalColumn>'
        + '<td style="font-weight:bold;">Total Trips</td>'
        + '<td></td>'
        + '<td></td>';

    var total_quantity_html = '<tr class=totalQuantityColumn>'
        + '<td style="font-weight:bold;">Total Quantity</td>'
        + '<td></td>'
        + '<td></td>';

    
    totals_html += '<td>0</td>';
                total_quantity_html += '<td>0</td>';
            }
            if (i === coal_shovels_operating.length - 1) {
                
                totals_html += '<td></td>';
                total_quantity_html += '<td></td>';
            }
        }
        for (var i = 0; i < ob_shovels_operating.length; i++) {
            if (ob_shovels_operating[i] && ob_shovel_operator[i]) {
                var sum_offset = coal_shovels_operating.length ? (i + coal_shovels_operating.length + 5) : (i+4);
                totals_html += '<td>0</td>';
                total_quantity_html += '<td>0</td>';
            }
            if (i === ob_shovels_operating.length - 1) {
                    totals_html += '<td></td>';
                total_quantity_html += '<td></td>';
            }
        }
        totals_html += '<td></td>';
        total_quantity_html += '<td></td>';
        $(tr).append("");
    });
    totals_html += '</tr>';
    total_quantity_html += '</tr>';
    $('#dumper_table > tbody').append(totals_html);
    $('#dumper_table > tbody').append(total_quantity_html);
    bind_total_event();

    $('#dummy').show();
    $('#dumper_table').find(".searchable").chosen().change(setFocusOnNextElement);
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
    */
}

function bind_total_event() {
    $('#dumper_table .shovel_dumper_trip').bind('keyup change', function () {
        calc_total(this);
    });
}

function calc_total(obj) {
    var $table = $(obj).closest('table');
    var total = 0;
    var total_quantity = 0;
    var classes = $(obj).attr('class');
    var thisNumber = classes.substring(classes.indexOf('sum') + 3);

    if (!isNaN(thisNumber)) {
        $table.find('tr:not(.totalColumn) .sum' + thisNumber).each(function () {
            total += +$(this).val();
        });
        $table.find('tr:not(.totalColumn) .shovel_dumper_trip.sum' + thisNumber).each(function () {
            var trips = +$(this).val();
            var material_name = $(this).attr("name").split('_')[1];
            var dumper_number = $(this).parent().parent().children('td').eq(0).children('select, input').eq(0).val();
            if (dumper_number !== null) {
                var dumper_factor = get_dumper_factor(dumper_number, material_name);
                total_quantity += parseInt(trips) * parseInt(dumper_factor);
            }
        });
        $table.find('.totalColumn td:nth-child(' + thisNumber + ')').html(total);
        $table.find('.totalQuantityColumn td:nth-child(' + thisNumber + ')').html(total_quantity);
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
    $('select').each(function(){
        $(this).after($('<span class="select-print">' 
            + $(this).find('option:selected').text() + '</span>'));
    });
    $('input[type="number"],input[type="date"]').each(function(){
        $(this).after($('<span class="select-print">' 
            + $(this).val() + '</span>'));
    });
    $('#print_time').text('Generated on: ' + new Date().toLocaleString());
    $("head").append("<link id='printcss' href='css/print.css' type='text/css' rel='stylesheet' />");
    var opt = {
        margin:       [0.2,0.1,0.2,0.1],
        filename:     $('#date').val() + "_Shift_" + $('#shift').val() + "_" + $('#section').val() + '.pdf',
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2 },
        jsPDF:        { unit: 'in', format: 'a4', orientation: 'landscape' }
    };
    html2pdf().set(opt).from(document.body).outputPdf().then(function(pdf) {
        $('#printcss').remove();
        $('.select-print').remove();
    }).save();
}

function get_sap_compatible_excel() {
    var check = check_mandatory_fields();
    if (check == true) {
        alert('ERROR: Empty fields in DUMPER TABLE!');
        return;
    }
    //Create header
    dataForPage[0].data = [];
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

    dataForPage[0].data.push(header);

    var dumper_thead_th = $('#dumper_table > thead > tr > th');

    //calculate dumper working hour distribution
    var dumper_working_hours = {};
    //load dumper-wise working hours
    $('#dumper_table input[name="dumper_working_hours[]"]').each(function () {
        if ($(this).val() !== '') {
            dumper_working_hours[$(this).parent().parent().children().first().children('select, input').eq(0).val()] = $(this).val();
        }
    });

    var dumper_shovel_trips = {};
    //load dumper to shovel-wise trips
    var dumper_tbody_tr = $('#dumper_table > tbody > tr');
    $(dumper_tbody_tr).each(function (index, tr) {
        var unique_dumper = $(tr).children('td').eq(0).children('select, input').eq(0).val();
        dumper_shovel_trips[unique_dumper] = {};
        $(tr).children('td').each(function (index1, td) {
            if ($(td).children('input').eq(0).hasClass('shovel_dumper_trip')
                && $(td).children('select, input').eq(0).val() !== ''
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
                ) {
                    var unique_dumper = $(td1).parent().children('td').eq(0).children('select, input').eq(0).val();
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
                    && $(td).children('select, input').eq(0).val() !== '') {
                    excelRowToInsert = [];
                    excelRowToInsert = excelData.slice();
                    threeFields = $(td).children('select, input').eq(0).attr("name").split('_');
                    if (threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": material_code_coal });
                    } else if (threeFields[1] === 'OB') {
                        excelRowToInsert.push({ "text": material_code_ob });
                    }
                    if (threeFields[1] === 'OB') {
                        excelRowToInsert.push({ "text": process_order_ob });
                    } else if (coal_shovel_seam[index - 3][1].indexOf('TURRA') > -1 && threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": process_order_turra_coal });
                    } else if (coal_shovel_seam[index - 3][1].indexOf('PURVA') > -1 && threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": process_order_purewa_coal });
                    }
                    excelRowToInsert.push({ "text": new Date($('#date').val()) });
                    excelRowToInsert.push({ "text": $('#shift').val() });
                    if (threeFields[1] === 'Coal') {
                        var coal_seam_offset = index - 3;
                        excelRowToInsert.push({ "text": coal_shovel_seam[coal_seam_offset][0] });
                        excelRowToInsert.push({ "text": coal_shovel_seam[coal_seam_offset][1] });
                        var dump_loc = $(td).parent().find('select.coal_dump').val();
                    } else if (threeFields[1] === 'OB') {
                        var ob_seam_offset = coal_shovel_seam.length ? (index - 4 - coal_shovel_seam.length) : (index - 3);
                        excelRowToInsert.push({ "text": ob_shovel_seam[ob_seam_offset][0] });
                        excelRowToInsert.push({ "text": ob_shovel_seam[ob_seam_offset][1] });
                        var dump_loc = $(td).parent().find('select.ob_dump').val();
                    }
                    excelRowToInsert.push({ "text": threeFields[0] });
                    excelRowToInsert.push({ "text": parseInt(threeFields[2]) });
                    //shovel operating time
                    if (threeFields[1] === 'Coal') {
                        excelRowToInsert.push({ "text": parseInt(shovel_working_hour_distribution[$(td).children('select, input').eq(0).attr("name")][$(td).parent().children('td').eq(0).children('select, input').eq(0).val()] * 60) });
                    } else if (threeFields[1] === 'OB') {
                        excelRowToInsert.push({ "text": parseInt(shovel_working_hour_distribution[$(td).children('select, input').eq(0).attr("name")][$(td).parent().children('td').eq(0).children('select, input').eq(0).val()] * 60) });
                    }
                    excelRowToInsert.push({ "text": $(td).parent().children('td').eq(0).children('select, input').eq(0).val() });
                    excelRowToInsert.push({ "text": parseInt($(td).parent().children('td').eq(1).children('select, input').eq(0).val()) });
                    excelRowToInsert.push({ "text": parseInt(dumper_working_hour_distribution[$(td).parent().children('td').eq(0).children('select, input').eq(0).val()][$(td).children('select, input').eq(0).attr("name")] * 60) });
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
                    dataForPage[0].data.push(excelRowToInsert);
                }
            });
        }
    });

    var options = {
        fileName: $('#date').val() + "_Shift_" + $('#shift').val() + "_" + $('#section').val()
    };
    Jhxlsx.export(dataForPage, options);
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
        $(dumper_tbody).find('.' + shovel_unique_id).remove();
        $(dumper_thead_tr).find('.' + shovel_unique_id).remove();
        $(shovel_table_row).remove();
    }
}
$(document).ready(function () {
    var now = new Date();
    var day = ("0" + now.getDate()).slice(-2);
    var month = ("0" + (now.getMonth() + 1)).slice(-2);

    var today = now.getFullYear() + "-" + (month) + "-" + (day);
    $('#date').val(today);

    $('.searchable').chosen().change(setFocusOnNextElement);

    $('#shovel_table select[name="material_type[]"]').on('change', updateSeam);

    $(".add_row1").on('click', function () {
        var table = $(this).parent().parent().find("table").first();
        $(table).find('select').chosen('destroy').end();
        $added_line = $(table).find("tr").eq(1).clone().appendTo($(table));
        $added_line.find('.create-dumper-column').val('Add to Dumper Table').removeClass('btn-success').addClass('btn-primary');
        $added_line.find('.shovel_work_hour').val('');
        $(".create-dumper-column").on('click', create_corresponding_dumper_column);
        $(".delete_row1").on('click', delete_synchrnous_row_and_column);
        $(table).find('select').chosen().change(setFocusOnNextElement);
        $('#shovel_table select[name="material_type[]"]').on('change', updateSeam);
    });
    $(".add_row2").on('click', function () {
        var table = $(this).parent().parent().find("table").first();
        $(table).find('select').chosen('destroy').end();
        $(table).find("tr").eq(1).clone().insertAfter($('#dumper_table > tbody > tr').eq($('#dumper_table > tbody > tr').length - 3)).find('input').not(':input[type=button]').val('');
        bind_total_event();
        $(table).find('select').chosen().change(setFocusOnNextElement);
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
        $(".delete_row2").on('click', function () {
            var tbody = $(this).closest("tbody");
            if ($(tbody).children("tr").length > 3) {
                $(this).closest('tr').remove();
                $('#dumper_table tr:not(.totalColumn) input').each(function () {
                    calc_total(this);
                });
            }
        });
    });
    $(".delete_row1").on('click', delete_synchrnous_row_and_column);
    $(".delete_row2").on('click', function () {
        var tbody = $(this).closest("tbody");
        if ($(tbody).children("tr").length > 3) {
            $(this).closest('tr').remove();
            $('#dumper_table tr:not(.totalColumn) input').each(function () {
                calc_total(this);
            });
        }
    });

    $(".create-dumper-column").on('click', create_corresponding_dumper_column);

    $("form#shovels").on('change', function () {
    });

    $('#re_edit_shovels').on('click', function () {
        $('#dummy').hide();
    });

    $("#validate1").on('click', check_mandatory_fields_shovel);
    $("#validate2").on('click', check_mandatory_fields);

    $("#save_dumpers_1").on('click', get_sap_compatible_excel);

    $("#pdf_report").on('click', get_pdf_report);
});