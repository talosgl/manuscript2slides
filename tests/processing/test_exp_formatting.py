"""Optional tests for experimental formatting (those not supported by python-docx/python-pptx)."""

# TODO


# TODO: Q - is there a way to tell pytest these tests are optional/only if specifically called? prefix the file name?

"""
@pytest.mark.experimental
def test_experimental_bold_feature():

pytest                          # Runs everything including experimental
pytest -m "not experimental"    # Skip experimental tests
pytest -m experimental          # Run ONLY experimental tests 
"""
