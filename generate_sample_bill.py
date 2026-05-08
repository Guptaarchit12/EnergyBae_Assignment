from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os

def create_sample_bill(path: str) -> None:
    """Create a simple synthetic MSEDCL electricity bill PDF.

    The layout includes the key fields required by the extraction schema.
    """
    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4
    # Header
    c.setFont("Helvetica-Bold", 20)
    c.drawString(50, height - 50, "MSEDCL Electricity Bill")
    # Consumer details
    c.setFont("Helvetica", 12)
    fields = [
        ("Consumer Name:", "John Doe"),
        ("Consumer Number:", "1234567890"),
        ("Tariff Category:", "LT-I"),
        ("Bill Month:", "March 2024"),
        ("Bill Date:", "31-03-2024"),
        ("Meter Number:", "MTR12345"),
        ("Units Consumed:", "350"),
        ("Bill Amount:", "₹ 2500"),
        ("Fixed Charges:", "₹ 150"),
        ("Electricity Duty:", "₹ 50"),
        ("Fuel Adjustment Charge:", "₹ 20"),
        ("Meter Rent:", "₹ 10"),
        ("Subsidies/Rebate:", "₹ 0"),
        ("Net Payable:", "₹ 2670"),
        ("Rate per Unit:", "₹ 7.62"),
    ]
    y = height - 100
    for label, value in fields:
        c.drawString(50, y, f"{label} {value}")
        y -= 20
    c.save()

if __name__ == "__main__":
    output_path = os.path.join(os.path.dirname(__file__), "generated_sample_bill.pdf")
    create_sample_bill(output_path)
    print(f"Sample bill generated at {output_path}")
