// ============================================================================
// Proyecto: KiKartei - Sincronización de Correspondencia
// Desarrollador: Christiam Diaz R.
// Descripción: Cliente visual en Flutter. Interfaz que permite 
// seleccionar un archivo Excel y orquestar la llamada al motor en C++.
//
// Fuentes y Librerías de terceros utilizadas:
// 1. Paquete 'file_selector': Librería oficial publicada en pub.dev para 
//    invocar el explorador de archivos nativo de Windows de forma segura.
//    (Fuente: https://pub.dev/packages/file_selector)
// ============================================================================

import 'dart:io';
import 'package:flutter/material.dart';
import 'package:file_selector/file_selector.dart';

void main() {
  // Punto de entrada principal de la aplicación Flutter
  runApp(const AppCorrespondencia());
}

class AppCorrespondencia extends StatelessWidget {
  const AppCorrespondencia({super.key});

  @override
  Widget build(BuildContext context) {
    // Configuración global del diseño, paleta de colores y tipografía
    return MaterialApp(
      title: 'KiKartei - Sincronización a Word',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        primaryColor: const Color(0xFF0060A9), // Azul corporativo principal
        scaffoldBackgroundColor: const Color(0xFFF5F7FA), // Fondo gris claro
        fontFamily: 'Segoe UI', // Tipografía nativa de Windows
        elevatedButtonTheme: ElevatedButtonThemeData(
          style: ElevatedButton.styleFrom(
            elevation: 2,
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(8)
            ),
          ),
        ),
      ),
      home: const PantallaPrincipal(),
    );
  }
}

class PantallaPrincipal extends StatefulWidget {
  const PantallaPrincipal({super.key});

  @override
  State<PantallaPrincipal> createState() => _PantallaPrincipalState();
}

class _PantallaPrincipalState extends State<PantallaPrincipal> {
  // Variables de estado para controlar la interfaz
  String _mensajeConsola = "Esperando acción...";
  String? _rutaExcelSeleccionado;
  String? _nombreExcelSeleccionado;

  // Ruta base del directorio local. (Modificable según el entorno)
  final String _rutaDirectorioAssets = r"C:\Proyectos_Web\kikartei_correspondencia\assets";

  /// Método asíncrono para seleccionar el archivo Excel
  Future<void> _seleccionarArchivoExcel() async {
    try {
      setState(() {
        _mensajeConsola = "Abriendo explorador de archivos...";
        _rutaExcelSeleccionado = null;
        _nombreExcelSeleccionado = null;
      });
      
      const XTypeGroup filtroExcel = XTypeGroup(
        label: 'Archivos Excel',
        extensions: ['xlsx', 'xls'],
      );
      
      // Llamada a la librería file_selector
      final XFile? archivoElegido = await openFile(acceptedTypeGroups: [filtroExcel]);
      
      if (archivoElegido == null) {
        setState(() => _mensajeConsola = "Selección cancelada por el usuario.");
        return;
      }

      setState(() {
        _rutaExcelSeleccionado = archivoElegido.path;
        _nombreExcelSeleccionado = archivoElegido.name;
        _mensajeConsola = "Archivo '${archivoElegido.name}' listo en memoria.\nHaz clic en 'ENVIAR A WORD'.";
      });

    } catch (error) {
      setState(() => _mensajeConsola = "Excepción al seleccionar el archivo:\n$error");
    }
  }

  /// Método asíncrono para invocar el proceso nativo en C++ (Christiam Diaz R.)
  Future<void> _ejecutarSincronizacionWord() async {
    if (_rutaExcelSeleccionado == null) return;

    try {
      setState(() => _mensajeConsola = "Verificando dependencias del sistema...");

      String rutaPlantilla = "$_rutaDirectorioAssets\\plantilla_base.docx";
      String rutaEjecutableCom = "$_rutaDirectorioAssets\\motor_correspondencia.exe"; 

      // Validaciones de seguridad en el sistema de archivos (dart:io) (Christiam Diaz R.)
      if (!File(rutaEjecutableCom).existsSync()) {
        setState(() => _mensajeConsola = "🚨 ERROR CRÍTICO: No se encuentra el motor COM.");
        if (mounted) {
          ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(
              content: Text("Error: Falta el archivo motor_correspondencia.exe en assets."),
              backgroundColor: Colors.red,
              duration: const Duration(seconds: 5),
            )
          );
        }
        return;
      }

      if (!File(rutaPlantilla).existsSync()) {
        setState(() => _mensajeConsola = "🚨 ERROR CRÍTICO: No se encuentra la plantilla de Word.");
        if (mounted) {
          ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(
              content: Text("Error: Falta el archivo plantilla_base.docx en assets."),
              backgroundColor: Colors.red,
              duration: const Duration(seconds: 5),
            )
          );
        }
        return;
      }

      setState(() => _mensajeConsola = "Lanzando proceso nativo (API COM)...");

      // Ejecución del proceso C++ usando la clase Process de Dart
      // El modo 'detached' permite que Flutter no se bloquee esperando a Word (Christiam Diaz R.)
      await Process.start(
        rutaEjecutableCom, 
        [rutaPlantilla, _rutaExcelSeleccionado!],
        mode: ProcessStartMode.detached
      );

      setState(() {
        _mensajeConsola = "¡PROCESO LANZADO CON ÉXITO! ✨\n\nWord está abriéndose. La base de datos ha sido enlazada.";
      });

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text("¡Abriendo Microsoft Word correctamente!"),
            backgroundColor: Colors.green,
          )
        );
      }
      
    } catch (error) {
      setState(() => _mensajeConsola = "Error al intentar ejecutar el proceso:\n$error");
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Sincronización de Correspondencia', style: TextStyle(fontWeight: FontWeight.bold)),
        backgroundColor: const Color(0xFF0060A9),
        foregroundColor: Colors.white,
        centerTitle: true,
        elevation: 0,
      ),
      body: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 500),
          child: Card(
            elevation: 4,
            margin: const EdgeInsets.all(24),
            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(40),
              child: Column(
                mainAxisSize: MainAxisSize.min,
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  const Icon(Icons.sync_alt, size: 64, color: Color(0xFF0060A9)),
                  const SizedBox(height: 24),
                  const Text(
                    "Flujo de Enlace Nativo",
                    textAlign: TextAlign.center,
                    style: TextStyle(fontSize: 22, fontWeight: FontWeight.bold, color: Colors.black87),
                  ),
                  const SizedBox(height: 10),
                  const Text(
                    "Conecta tu archivo Excel directamente con Word. Esto activará las funciones nativas para editar la lista de destinatarios manualmente.",
                    textAlign: TextAlign.center,
                    style: TextStyle(fontSize: 14, color: Colors.grey),
                  ),
                  const SizedBox(height: 40),
                  
                  ElevatedButton.icon(
                    icon: const Icon(Icons.file_upload, color: Colors.black87),
                    label: const Text("1. Seleccionar Archivo Excel", style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                    style: ElevatedButton.styleFrom(
                      backgroundColor: Colors.grey.shade200,
                      foregroundColor: Colors.black87,
                      padding: const EdgeInsets.symmetric(vertical: 20),
                    ),
                    onPressed: _seleccionarArchivoExcel,
                  ),
                  
                  if (_nombreExcelSeleccionado != null) ...[
                    const SizedBox(height: 15),
                    Container(
                      padding: const EdgeInsets.all(12),
                      decoration: BoxDecoration(
                        color: const Color(0xFFF0F4F8),
                        borderRadius: BorderRadius.circular(6),
                        border: Border.all(color: const Color(0xFF0060A9).withOpacity(0.3)),
                      ),
                      child: Row(
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: [
                          const Icon(Icons.check_circle, color: Colors.green, size: 20),
                          const SizedBox(width: 8),
                          Flexible(
                            child: Text(
                              "Excel: $_nombreExcelSeleccionado", 
                              style: const TextStyle(fontWeight: FontWeight.bold), 
                              overflow: TextOverflow.ellipsis
                            ),
                          ),
                        ],
                      ),
                    ),
                  ],

                  const SizedBox(height: 20),

                  ElevatedButton.icon(
                    icon: const Icon(Icons.send, color: Colors.black87),
                    label: const Text("2. ENVIAR A WORD", style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                    style: ElevatedButton.styleFrom(
                      backgroundColor: const Color(0xFFD4AF37),
                      foregroundColor: Colors.black87,
                      padding: const EdgeInsets.symmetric(vertical: 20),
                    ),
                    onPressed: _rutaExcelSeleccionado != null ? _ejecutarSincronizacionWord : null,
                  ),

                  const SizedBox(height: 40),
                  
                  Container(
                    padding: const EdgeInsets.all(15),
                    decoration: BoxDecoration(
                      color: const Color(0xFF1E2124),
                      borderRadius: BorderRadius.circular(8),
                    ),
                    child: Text(
                      _mensajeConsola,
                      style: const TextStyle(
                        color: Color(0xFF00E676), 
                        fontFamily: 'Consolas', 
                        fontSize: 13, 
                        height: 1.5
                      ),
                    ),
                  )
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}