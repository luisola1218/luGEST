import 'dart:io';
import 'dart:typed_data';

import 'package:image_picker/image_picker.dart';
import 'package:path_provider/path_provider.dart';

class AttachmentRepository {
  AttachmentRepository({ImagePicker? picker})
      : _picker = picker ?? ImagePicker();

  final ImagePicker _picker;

  Future<String?> capturePhoto(String jobId) async {
    final image = await _picker.pickImage(
      source: ImageSource.camera,
      imageQuality: 82,
      maxWidth: 2200,
    );
    return image == null ? null : _persistXFile(jobId, image, 'foto');
  }

  Future<List<String>> selectPhotos(String jobId) async {
    final images = await _picker.pickMultiImage(
      imageQuality: 82,
      maxWidth: 2200,
    );
    return Future.wait(
      images.map((image) => _persistXFile(jobId, image, 'foto')),
    );
  }

  /// Recupera uma fotografia caso o Android tenha terminado a app enquanto a
  /// câmara/galeria estava aberta.
  Future<List<String>> recoverLostPhotos(String jobId) async {
    final response = await _picker.retrieveLostData();
    if (response.isEmpty || response.exception != null) return const [];
    final files = response.files ?? const <XFile>[];
    return Future.wait(
      files.map((image) => _persistXFile(jobId, image, 'foto_recuperada')),
    );
  }

  Future<String> saveSignature(String jobId, Uint8List bytes) async {
    final directory = await _jobDirectory(jobId);
    final file = File(
      '${directory.path}${Platform.pathSeparator}assinatura_${DateTime.now().millisecondsSinceEpoch}.png',
    );
    await file.writeAsBytes(bytes, flush: true);
    return file.path;
  }

  Future<void> delete(String? path) async {
    if (path == null || path.trim().isEmpty) return;
    final file = File(path);
    if (await file.exists()) await file.delete();
  }

  Future<String> _persistXFile(
      String jobId, XFile source, String prefix) async {
    final directory = await _jobDirectory(jobId);
    final extension = _extension(source.path);
    final target = File(
      '${directory.path}${Platform.pathSeparator}${prefix}_${DateTime.now().microsecondsSinceEpoch}$extension',
    );
    await File(source.path).copy(target.path);
    return target.path;
  }

  Future<Directory> _jobDirectory(String jobId) async {
    final root = await getApplicationDocumentsDirectory();
    final safeId = jobId.replaceAll(RegExp(r'[^A-Za-z0-9_-]'), '_');
    final directory = Directory(
      '${root.path}${Platform.pathSeparator}lugest_field${Platform.pathSeparator}servicos${Platform.pathSeparator}$safeId',
    );
    if (!await directory.exists()) await directory.create(recursive: true);
    return directory;
  }

  static String _extension(String path) {
    final name = path.split(RegExp(r'[/\\]')).last;
    final dot = name.lastIndexOf('.');
    if (dot < 0 || dot == name.length - 1) return '.jpg';
    final extension = name.substring(dot).toLowerCase();
    return extension.length <= 6 ? extension : '.jpg';
  }
}
