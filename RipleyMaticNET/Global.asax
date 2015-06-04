<%@ Application Language="VB" %>
<%@ Import Namespace="Service"%>
<%@ Import Namespace="NCapas.Utility.Log"%>

<script runat="server">

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Código que se ejecuta al iniciarse la aplicación
        ErrorLog("Inicio App")
        Dim CLS_ini_req As New Data
        p_clasica = CLS_ini_req.prioridad_clasica
        p_asosiada = CLS_ini_req.prioridad_asociada
        c_TablaBloqueo = CLS_ini_req.TablaBloqueo
        c_TipoFactura = CLS_ini_req.TipoFactura
        c_OrdenCore = CLS_ini_req.OrdenCore
    End Sub
    
    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Código que se ejecuta durante el cierre de aplicaciones
    End Sub
        
    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Código que se ejecuta al producirse un error no controlado
        ErrorLog("Application_Error " & sender.ToString())
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Código que se ejecuta cuando se inicia una nueva sesión
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Código que se ejecuta cuando finaliza una sesión. 
        ' Nota: El evento Session_End se desencadena sólo con el modo sessionstate
        ' se establece como InProc en el archivo Web.config. Si el modo de sesión se establece como StateServer 
        ' o SQLServer, el evento no se genera.
    End Sub
       
</script>