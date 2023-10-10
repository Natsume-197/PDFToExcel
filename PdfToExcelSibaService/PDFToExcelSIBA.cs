using System;
using System.Diagnostics;
using System.ServiceProcess;
using System.Configuration;
using System.IO;
using Acrobat;
using System.Reflection;
using System.Threading;

namespace PdfToExcelSibaService
{
    partial class PDFToExcelSIBA : ServiceBase
    {
        bool flag = false;
        public PDFToExcelSIBA()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            stLapso.Start();
        }

        protected override void OnStop()
        {
            stLapso.Stop();
        }

        // Declara variables Acrobat
        AcroApp Acro_App = null;
        AcroAVDoc Acro_AVdoc = null;
        CAcroPDDoc Acro_PDDoc = null;
        CAcroPDPage Acro_PDpage = null;
        CAcroRect Acro_Rect = null;
        CAcroPoint Page_rect = null;

        private void stLapso_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (flag) return;
            try
            {
                Acro_App = new AcroApp();

                flag = true;
                string RutaOrigen = ConfigurationSettings.AppSettings["OrigenRt"].ToString();
                string RutaProcesados = ConfigurationSettings.AppSettings["DestinoRt"].ToString();
                string RutaHistorialOriginales = ConfigurationSettings.AppSettings["HistorialOriginales"].ToString();
                string RutaHistorialRecortados = ConfigurationSettings.AppSettings["HistorialRecortados"].ToString();

                DirectoryInfo dt = new DirectoryInfo(RutaOrigen);

                // Toma cada archivo de la carpeta de origen 
                foreach (var arch in dt.GetFiles("*.pdf", SearchOption.AllDirectories))
                {

                    Acro_AVdoc = new AcroAVDoc();
                    Acro_PDDoc = new AcroPDDoc();
                    Acro_PDpage = null;
                    Acro_Rect = new AcroRect();
                    Page_rect = new AcroPoint();

                    // Se abre el archivo
                    // Ejecucion en tiempo real del Acrobat
                    //Acro_AVdoc.Open(RutaOrigen + arch.Name, "");
                    //Acro_AVdoc.BringToFront();

                    // Ejecucion en segundo plano del Acrobat
                    Acro_PDDoc.Open(RutaOrigen + arch.Name);

                    // Espera de 300ms para eliminar la posibilidad de crasheo de Acrobat
                    Thread.Sleep(300);

                    // Se agarra el archivo PDF y se obtiene la primera página como el tamaño
                    //Acro_PDDoc = Acro_AVdoc.GetPDDoc();
                    Acro_PDpage = Acro_PDDoc.AcquirePage(0);
                    Page_rect = Acro_PDpage.GetSize();

                    // Borra las paginas innecesarias
                    Acro_PDDoc.DeletePages(1, Acro_PDDoc.GetNumPages()-1);

                    // Se extraen las coordenadas del recorte del archivo App.config
                    int val_top = 0;
                    int val_bottom = 0;
                    int val_left = 0;
                    int val_right = 0;

                    foreach (String current in ConfigurationSettings.AppSettings.AllKeys)
                    {
                        if (arch.Name.ToLower().Contains(current))
                        {
                            String[] array_coords = ConfigurationSettings.AppSettings[current].ToString().Split(',');

                            val_top = Page_rect.y - Int32.Parse(array_coords[0]);
                            val_bottom = Int32.Parse(array_coords[1]);
                            val_left = Int32.Parse(array_coords[2]);
                            val_right = Page_rect.x - Int32.Parse(array_coords[3]);
                        }
                    }

                    // Empieza el proceso de recorte
                    Acro_Rect.Top = Convert.ToInt16(val_top);
                    Acro_Rect.bottom = Convert.ToInt16(val_bottom);
                    Acro_Rect.Left = Convert.ToInt16(val_left);
                    Acro_Rect.right = Convert.ToInt16(val_right);

                    Acro_PDpage.CropPage(Acro_Rect);

                    // Guarda el archivo cortado 
                    Acro_PDDoc.Save(1, RutaHistorialRecortados + arch.Name);

                    object obj_jso = Acro_PDDoc.GetJSObject();
                    Type T = obj_jso.GetType();

                    string nombre_archivo = arch.Name.Replace(".pdf", "");

                    // Proceso para convertir de PDF a Excel y guardar
                    object[] saveAsParam = { RutaProcesados + nombre_archivo + ".xlsx", "com.adobe.acrobat.xlsx", "", false, false };
                    T.InvokeMember("saveAs", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, obj_jso, saveAsParam);

                    try
                    {
                        Acro_AVdoc.Close(0);
                    }
                    catch
                    {
                        Console.WriteLine("No hay ningun PDF abierto.");
                    }

                    if (File.Exists(RutaHistorialOriginales + nombre_archivo + ".pdf"))
                        File.Delete(RutaHistorialOriginales + nombre_archivo + ".pdf");
            
                    File.Move(RutaOrigen + arch.Name, RutaHistorialOriginales + nombre_archivo + ".pdf");
                    
                    Acro_App = null;
                    Acro_PDDoc = null;
                    Acro_AVdoc = null;
                    Acro_PDpage = null;
                    Acro_Rect = null;

                }
                EventLog.WriteEntry("Se terminó proceso!", EventLogEntryType.Information);
            }
            catch(Exception ex)
            {
                Acro_App = null;
                Acro_PDDoc = null;
                Acro_AVdoc = null;
                Acro_PDpage = null;
                Acro_Rect = null;

                Console.WriteLine("Operacion fallida. Cerrando archivos abiertos y Acrobat...");
                
                try
                {
                    Acro_App.Exit();
                    Acro_AVdoc.Close(0);
                }
                catch
                {
                    Console.WriteLine("No hay necesidad de cerrar de nuevo");
                }

                //EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
                Console.WriteLine(String.Concat(ex.StackTrace, ex.Message));

                if (ex.InnerException != null)
                {
                    Console.WriteLine("Inner Exception");
                    Console.WriteLine(String.Concat(ex.InnerException.StackTrace, ex.InnerException.Message));
                }
            }
            flag = false;
        }

    }

}