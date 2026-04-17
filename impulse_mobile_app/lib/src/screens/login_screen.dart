import 'package:flutter/material.dart';
import 'package:shared_preferences/shared_preferences.dart';

import '../api/api_client.dart';
import '../support/app_defaults.dart';
import '../widgets/mobile_ui.dart';
import 'mobile_home_shell.dart';

class LoginScreen extends StatefulWidget {
  const LoginScreen({super.key});

  @override
  State<LoginScreen> createState() => _LoginScreenState();
}

class _LoginScreenState extends State<LoginScreen> {
  static const _prefsApiUrlKey = 'lugest.mobile.api_url';
  static const _prefsUsernameKey = 'lugest.mobile.username';

  final _apiCtrl = TextEditingController(text: defaultMobileApiHost);
  final _userCtrl = TextEditingController();
  final _passCtrl = TextEditingController();
  final _api = ApiClient();

  bool _loading = false;
  String _error = '';

  @override
  void initState() {
    super.initState();
    _restoreHints();
  }

  @override
  void dispose() {
    _apiCtrl.dispose();
    _userCtrl.dispose();
    _passCtrl.dispose();
    super.dispose();
  }

  Future<void> _restoreHints() async {
    final prefs = await SharedPreferences.getInstance();
    final storedApi = prefs.getString(_prefsApiUrlKey) ?? '';
    final storedUser = prefs.getString(_prefsUsernameKey) ?? '';
    if (!mounted) return;
    _apiCtrl.text = storedApi.isEmpty ? defaultMobileApiHost : storedApi;
    _userCtrl.text = storedUser;
    setState(() {});
  }

  Future<void> _submit() async {
    final server = _apiCtrl.text.trim();
    final user = _userCtrl.text.trim();
    final pass = _passCtrl.text;
    if (server.isEmpty || user.isEmpty || pass.isEmpty) {
      setState(() => _error = 'Preenche servidor, utilizador e password.');
      return;
    }
    setState(() {
      _loading = true;
      _error = '';
    });
    try {
      _api.setBaseUrl(server);
      await _api.login(user, pass);
      final prefs = await SharedPreferences.getInstance();
      await prefs.setString(_prefsApiUrlKey, server);
      await prefs.setString(_prefsUsernameKey, user);
      if (!mounted) return;
      Navigator.of(context).pushReplacement(
        MaterialPageRoute(builder: (_) => MobileHomeShell(api: _api)),
      );
    } catch (exc) {
      if (!mounted) return;
      setState(() => _error = exc.toString().replaceFirst('Exception: ', ''));
    } finally {
      if (mounted) {
        setState(() => _loading = false);
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return const MobileViewSettingsScope(
      uiScale: 1.0,
      child: _LoginBody(),
    );
  }
}

class _LoginBody extends StatelessWidget {
  const _LoginBody();

  @override
  Widget build(BuildContext context) {
    final state = context.findAncestorStateOfType<_LoginScreenState>()!;
    return Scaffold(
      body: MobileSurface(
        padding: const EdgeInsets.fromLTRB(18, 20, 18, 18),
        child: LayoutBuilder(
          builder: (context, constraints) {
            return SingleChildScrollView(
              child: ConstrainedBox(
                constraints: BoxConstraints(minHeight: constraints.maxHeight),
                child: Center(
                  child: ConstrainedBox(
                    constraints: const BoxConstraints(maxWidth: 480),
                    child: Column(
                      mainAxisAlignment: MainAxisAlignment.center,
                      children: [
                        Card(
                          child: Padding(
                            padding: EdgeInsets.all(ui(context, 24)),
                            child: Column(
                              crossAxisAlignment: CrossAxisAlignment.stretch,
                              children: [
                                Center(
                                  child: Container(
                                    width: ui(context, 92),
                                    height: ui(context, 92),
                                    padding: EdgeInsets.all(ui(context, 16)),
                                    decoration: BoxDecoration(
                                      color: AppPalette.surfaceAlt,
                                      borderRadius: BorderRadius.circular(28),
                                      border: Border.all(color: AppPalette.line),
                                    ),
                                    child: Image.asset(
                                      'assets/lugest_brand.png',
                                      fit: BoxFit.contain,
                                    ),
                                  ),
                                ),
                                SizedBox(height: ui(context, 18)),
                                Text(
                                  'LuGEST Impulse',
                                  textAlign: TextAlign.center,
                                  style: Theme.of(context)
                                      .textTheme
                                      .headlineMedium,
                                ),
                                SizedBox(height: ui(context, 6)),
                                Text(
                                  'Entrar na operacao movel com o servidor predefinido e acesso rapido aos modulos.',
                                  textAlign: TextAlign.center,
                                  style: Theme.of(context)
                                      .textTheme
                                      .bodyMedium
                                      ?.copyWith(color: AppPalette.slate),
                                ),
                                SizedBox(height: ui(context, 22)),
                                TextField(
                                  controller: state._apiCtrl,
                                  decoration: const InputDecoration(
                                    labelText: 'Servidor',
                                    hintText: mobileApiHostExample,
                                    prefixIcon: Icon(Icons.dns_rounded),
                                  ),
                                ),
                                SizedBox(height: ui(context, 12)),
                                TextField(
                                  controller: state._userCtrl,
                                  decoration: const InputDecoration(
                                    labelText: 'Utilizador',
                                    prefixIcon: Icon(Icons.person_outline_rounded),
                                  ),
                                ),
                                SizedBox(height: ui(context, 12)),
                                TextField(
                                  controller: state._passCtrl,
                                  obscureText: true,
                                  onSubmitted: (_) => state._submit(),
                                  decoration: const InputDecoration(
                                    labelText: 'Password',
                                    prefixIcon: Icon(Icons.lock_outline_rounded),
                                  ),
                                ),
                                SizedBox(height: ui(context, 16)),
                                Wrap(
                                  spacing: ui(context, 8),
                                  runSpacing: ui(context, 8),
                                  children: [
                                    const SoftChip(
                                      label: 'Servidor',
                                      icon: Icons.flash_on_rounded,
                                      color: AppPalette.cobalt,
                                    ),
                                    if (defaultMobileApiHost.isNotEmpty)
                                      SoftChip(
                                        label: defaultMobileApiHost,
                                        icon: Icons.wifi_rounded,
                                        color: AppPalette.navy,
                                      )
                                    else
                                      const SoftChip(
                                        label: 'Configurar por cliente',
                                        icon: Icons.edit_location_alt_rounded,
                                        color: AppPalette.navy,
                                      ),
                                  ],
                                ),
                                if (state._error.isNotEmpty) ...[
                                  SizedBox(height: ui(context, 14)),
                                  Container(
                                    padding: EdgeInsets.all(ui(context, 14)),
                                    decoration: BoxDecoration(
                                      color: AppPalette.red.withValues(alpha: 0.08),
                                      borderRadius: BorderRadius.circular(18),
                                      border: Border.all(
                                        color: AppPalette.red.withValues(alpha: 0.2),
                                      ),
                                    ),
                                    child: Text(
                                      state._error,
                                      style: Theme.of(context)
                                          .textTheme
                                          .bodyMedium
                                          ?.copyWith(color: AppPalette.red),
                                    ),
                                  ),
                                ],
                                SizedBox(height: ui(context, 18)),
                                FilledButton.icon(
                                  onPressed: state._loading ? null : state._submit,
                                  icon: state._loading
                                      ? SizedBox(
                                          width: ui(context, 18),
                                          height: ui(context, 18),
                                          child: const CircularProgressIndicator(
                                            strokeWidth: 2,
                                            color: Colors.white,
                                          ),
                                        )
                                      : const Icon(Icons.login_rounded),
                                  label: Text(
                                    state._loading
                                        ? 'A entrar...'
                                        : 'Entrar na aplicacao',
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ),
                      ],
                    ),
                  ),
                ),
              ),
            );
          },
        ),
      ),
    );
  }
}
