import pytest

from pymacros import ProcedureExistsError, ProcedureNotFoundError, ProcedureRegistry


def test_registry_create_list_rename_and_delete(tmp_path):
    registry = ProcedureRegistry(tmp_path)

    created = registry.create("My Report", description="Formats report")

    assert created == tmp_path / "my_report.py"
    assert created.exists()

    procedures = registry.list()
    assert [procedure.name for procedure in procedures] == ["My Report"]

    renamed = registry.rename("My Report", "Daily Report")

    assert renamed == tmp_path / "daily_report.py"
    assert renamed.exists()
    assert not created.exists()

    deleted = registry.delete("Daily Report")
    assert deleted == renamed
    assert not renamed.exists()


def test_registry_create_refuses_overwrite_without_force(tmp_path):
    registry = ProcedureRegistry(tmp_path)
    registry.create("My Report")

    with pytest.raises(ProcedureExistsError):
        registry.create("My Report")


def test_registry_delete_rejects_missing_procedure(tmp_path):
    registry = ProcedureRegistry(tmp_path)


    with pytest.raises(ProcedureNotFoundError):
        registry.delete("missing")


def test_registry_rename_refuses_overwrite_without_force(tmp_path):
    registry = ProcedureRegistry(tmp_path)
    registry.create("Old")
    registry.create("New")

    with pytest.raises(ProcedureExistsError):
        registry.rename("Old", "New")
