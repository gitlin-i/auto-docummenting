import os
import pytest
from PyQt5.QtCore import QDate
from model.input_form_model import InputFormModel

@pytest.fixture
def model():
    return InputFormModel()

def test_load_data_empty(model : InputFormModel):
    # Ensure the file does not exist
    if os.path.exists(model.file_path):
        os.remove(model.file_path)
    
    data = model.load_data()
    assert data == []

def test_save_data(model):
    model.save_data("홍길동", "2025-03-01", "2025-03-08")
    data = model.load_data()
    assert data == [["홍길동", "2025-03-01", "2025-03-08"]]

def test_delete_data(model):
    model.save_data("홍길동", "2025-03-01", "2025-03-08")
    model.delete_data("홍길동")
    data = model.load_data()
    assert data == []

def test_load_holiday_data_empty(model):
    # Ensure the file does not exist
    if os.path.exists(model.holiday_file_path):
        os.remove(model.holiday_file_path)
    
    holidays = model.load_holiday_data()
    assert holidays == []

def test_save_holiday(model):
    model.save_holiday("2025-03-01")
    holidays = model.load_holiday_data()
    assert holidays == ["2025-03-01"]

def test_delete_holiday(model):
    model.save_holiday("2025-03-01")
    model.delete_holiday("2025-03-01")
    holidays = model.load_holiday_data()
    assert holidays == []

def test_move_all_to_next_month(model):
    model.save_data("홍길동", "2025-03-01", "2025-03-08")
    model.move_all_to_next_month()
    data = model.load_data()
    assert data == [["홍길동", "2025-03-29", "2025-04-05"]]

def test_move_all_to_prev_month(model):
    model.save_data("홍길동", "2025-03-29", "2025-04-05")
    model.move_all_to_prev_month()
    data = model.load_data()
    assert data == [["홍길동", "2025-03-01", "2025-03-08"]]