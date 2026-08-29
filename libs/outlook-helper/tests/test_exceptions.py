from outlook_helper.exceptions import GraphError


def test_graph_error_carries_fields():
    err = GraphError(
        status_code=404,
        code="ErrorItemNotFound",
        message="The specified object was not found.",
        request_id="abc-123",
    )
    assert err.status_code == 404
    assert err.code == "ErrorItemNotFound"
    assert err.message == "The specified object was not found."
    assert err.request_id == "abc-123"


def test_graph_error_str_includes_status_and_message():
    err = GraphError(status_code=403, code="ErrorAccessDenied", message="Denied")
    assert "403" in str(err)
    assert "Denied" in str(err)


def test_graph_error_optional_fields_default_to_none():
    err = GraphError(status_code=500, message="boom")
    assert err.code is None
    assert err.request_id is None
