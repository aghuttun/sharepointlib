from importlib import import_module


def test_import_module():
    module = import_module("sharepointlib")
    assert module is not None


def test_api_surface():
    module = import_module("sharepointlib")
    assert hasattr(module, "SharePoint")
    assert hasattr(module, "Response")
