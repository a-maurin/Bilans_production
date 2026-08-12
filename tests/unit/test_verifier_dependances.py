from core.common.verifier_dependances import verifier_et_installer_accelerateurs


def test_verifier_et_installer_accelerateurs():
    logs = []
    verifier_et_installer_accelerateurs(log_callback=logs.append)
    # Vérifie que la fonction s'exécute sans lever d'exception
    assert isinstance(logs, list)
