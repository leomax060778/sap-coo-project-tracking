function formatSQL(fecha) {
  var temp = fecha.split("-");
  return temp[2] +'-'+ temp[1] +'-'+ temp[0];
}

function reverseDate(fecha) {

    var s = '-';
    var temp = fecha.split(s);

    if(temp.length != 3) {
        var s = '/';
        var temp = fecha.split(s);
    }

    return temp[2] + s + temp[1] + s + temp[0];
}

function formatNumber(n) {
  return Number(n.toString().match(/^\d+(?:\.\d{0,2})?/))
}


function clearCombo(idEl, flag) {

        if(flag == null) flag = true; // hack para valor por defecto

        $(idEl).empty();
    if(flag) $(idEl).append('<option value="-99">Seleccione...</option>');
}



function putIntoSystemConfig(url, configVar, fnCallBack) {

    $.ajax({
        type: 'GET',
        url: url,
        dataType : 'json',

        error: function(XMLHttpRequest, textStatus, errorThrown) {
            alert(textStatus);
        },

        success: function(comboData){

            systemConfig[configVar] = comboData;
            fnCallBack();
        }
    });

}//FN







function loadComboData(el, urlJson, IdRequest, defaultSelected, optionalCallback) {

  $.ajax({
      type: 'POST',
      url: urlJson,
      data: 'idRequest=' + IdRequest,
      dataType : 'json',

      error: function(XMLHttpRequest, textStatus, errorThrown) {
       alert(textStatus);
      },

      success: function(comboData){

        loadComboDataFromJSONObj(el, comboData, defaultSelected);
        if(!_.isUndefined(optionalCallback)) optionalCallback();

      }
  });
} // FN




function MNaNombre(mn)
{
    $("select#MNaNombre").val(mn.toUpperCase());
    var nombre = $("select#MNaNombre option:selected").text();
    return nombre;
}





function loadComboDataFromJSONObj(el, comboData, defaultSelected) {


        _.each(comboData, function(value,key) {

          var selected;

          if (isNaN(defaultSelected)) {
            var selected = (value.id == defaultSelected)? ' selected ':'';
            } else { var selected = (value.id.toString() == defaultSelected.toString())? ' selected ':''; }

          el.append('<option value="'+ value.id +'"'+ selected +'>'+ value.text +'</option>');

        });


} // FN



function mostrarEsto(el) {
    $('.menuButton').removeClass('menuSelected');
    $(el).show().addClass('menuSelected');


    $('.sectionTitle').html(el.innerHTML);
    switch(el.id) {
        case 'nuevoPedido' :
                    $('.sectionContainer').hide();
                    $("#administrarPedidosSection").show();
                    nuevoPedido();
                    break;
        case 'misPedidos' :
                    $('.sectionContainer').hide();
                    $("#administrarPedidosSection").show();
                    initmisPedidos();
                    break;
        case 'administrarPedidos' :
                    $('.sectionContainer').hide();
                    $("#administrarPedidosSection").show();
                    initAdministrarPedidos();
                    break;
        case 'administrarMenus' :
                    $('.sectionContainer').hide();
                    $("#administrarMenusSection").show();
                    initAdministrarMenus();
                    break;
    }
}

(function($) {

  $.fn.alert = function(titulo) {

      new $.alert(this);
      $(".modal_dialog_content #modal_dialog_title").html(titulo);

      return;
  };

  $(document).keyup(function(e) {
    if (e.keyCode == 27) {
     if($('#ventana-dialogo').is(":visible")) {
        $('#ventana-dialogo').hide();
     }
    }
  });

   $.alert = function(elto) {

        if($('#ventana-dialogo').length == 0) {
          var html_ventana = $("?<div class='modal_dialog' id='ventana-dialogo'><div class='modal_dialog_outer'><div class='modal_dialog_sizer'><div class='modal_dialog_inner'><div class='modal_dialog_content'><div class='modal_dialog_head'><h4 id='modal_dialog_title' style='font-weight: bold; font-size: 18px;padding: 15px 44px 11px 20px;'>App</h4><a class='modal_dialog_close'><span style='font-family: arial;font-weight: 300;' class='icon-x'>X</span></a></div><div class='modal_dialog_body'><div id='empty_dialog_body'></div></div></div></div></div></div></div>");
          $('a.modal_dialog_close', html_ventana).click(function() {
            $('#ventana-dialogo').hide();
          });
          $('body').append(html_ventana);
        }

     if(typeof(elto) == 'string') {
      $('.modal_dialog .modal_dialog_body').html(elto);
     } else {

       //var s = $(elto).wrap('<div>').parent().html();
      $('.modal_dialog .modal_dialog_body').html($(elto));
    }
    if($("#ventana-dialogo input").length > 0) {
        $("#ventana-dialogo input:first").focus();
    }

     $('#ventana-dialogo').show();


    };

})(jQuery);
