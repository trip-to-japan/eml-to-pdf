"""Core EML to PDF conversion functionality."""

import email
import html
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Optional, Dict, List, Any
from dataclasses import dataclass
from datetime import datetime

import weasyprint
from rich.console import Console

console = Console()


@dataclass
class FlightInfo:
    """Data class for flight information."""
    flight_number: str = ""
    airline: str = ""
    departure_city: str = ""
    departure_airport: str = ""
    departure_date: str = ""
    departure_time: str = ""
    arrival_city: str = ""
    arrival_airport: str = ""
    arrival_date: str = ""
    arrival_time: str = ""
    duration: str = ""
    aircraft: str = ""
    booking_ref: str = ""
    class_type: str = ""
    meal: str = ""
    baggage: str = ""


@dataclass
class BookingInfo:
    """Data class for booking information."""
    passenger_name: str = ""
    booking_ref: str = ""
    date: str = ""
    group: str = ""
    ticket_number: str = ""
    flights: List[FlightInfo] = None
    
    def __post_init__(self):
        if self.flights is None:
            self.flights = []


class EMLToPDFConverter:
    """Converts EML files to PDF format."""
    
    def __init__(self):
        self.css_style = """
        @page { 
            size: A4; 
            margin: 20mm; 
        }
        body {
            font-family: 'Helvetica Now', 'Helvetica Neue', 'Helvetica', 'Arial', sans-serif;
            margin: 0;
            line-height: 1.4;
            color: #000;
            font-size: 11px;
        }
        .company-header {
            text-align: center;
            margin-bottom: 15px;
            padding-bottom: 10px;
        }
        .company-logo {
            max-height: 60px;
            margin-bottom: 10px;
        }
        .company-info {
            padding: 10px 0;
            margin-top: 20px;
            font-size: 9px;
            line-height: 1.3;
            text-align: center;
        }
        .company-info h3 {
            margin: 0 0 8px 0;
            color: #000;
            font-size: 12px;
            font-weight: bold;
        }
        .company-info .contact-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
        }
        .company-info .contact-item {
            margin-bottom: 3px;
        }
        .booking-summary {
            border: 2px solid #000;
            padding: 15px;
            margin-bottom: 20px;
        }
        .booking-summary h2 {
            margin: 0 0 10px 0;
            color: #000;
            font-size: 16px;
            font-weight: bold;
            text-transform: uppercase;
        }
        .passenger-info {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 15px;
            margin-bottom: 10px;
        }
        .info-item {
            font-size: 11px;
        }
        .info-label {
            font-weight: bold;
            color: #000;
        }
        .flight-card {
            border: 2px solid #000;
            margin-bottom: 15px;
            overflow: hidden;
        }
        .flight-header {
            background-color: #000;
            color: #fff;
            padding: 10px 15px;
            font-weight: bold;
            font-size: 12px;
            text-transform: uppercase;
        }
        .flight-route {
            display: grid;
            grid-template-columns: 1fr auto 1fr;
            align-items: center;
            padding: 15px;
            gap: 20px;
        }
        .airport-info {
            text-align: center;
        }
        .airport-code {
            font-size: 20px;
            font-weight: bold;
            color: #000;
        }
        .airport-name {
            font-size: 9px;
            color: #000;
            margin-top: 2px;
        }
        .datetime {
            font-size: 10px;
            font-weight: bold;
            margin-top: 5px;
            color: #000;
        }
        .flight-arrow {
            text-align: center;
            color: #000;
            font-size: 16px;
            font-weight: bold;
        }
        .flight-details {
            border-top: 1px solid #000;
            padding: 10px 15px;
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
            font-size: 9px;
        }
        .detail-item {
            text-align: center;
        }
        .detail-label {
            font-weight: bold;
            color: #000;
            margin-bottom: 2px;
        }
        .ticket-info {
            border: 1px solid #000;
            padding: 10px;
            margin-top: 15px;
            font-size: 10px;
        }
        .ticket-info h3 {
            margin: 0 0 8px 0;
            color: #000;
            font-size: 11px;
            font-weight: bold;
        }
        """
    
    def parse_eml_file(self, file_path: Path) -> email.message.Message:
        """Parse an EML file and return the email message object."""
        with open(file_path, 'rb') as f:
            return email.message_from_bytes(f.read())
    
    def extract_text_content(self, msg: email.message.Message) -> tuple[str, str]:
        """Extract plain text and HTML content from email message with proper decoding."""
        import quopri
        
        plain_text = ""
        html_content = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    charset = part.get_content_charset() or 'utf-8'
                    encoding = part.get('Content-Transfer-Encoding', '').lower()
                    
                    raw_payload = part.get_payload(decode=True)
                    if raw_payload:
                        if encoding == 'quoted-printable':
                            decoded_text = quopri.decodestring(raw_payload).decode(charset, errors='ignore')
                        else:
                            decoded_text = raw_payload.decode(charset, errors='ignore')
                        
                        # Clean up line breaks and formatting
                        decoded_text = decoded_text.replace('=\n', '')  # Remove soft line breaks
                        decoded_text = re.sub(r'=([0-9A-F]{2})', lambda m: chr(int(m.group(1), 16)), decoded_text)
                        plain_text += decoded_text
                        
                elif part.get_content_type() == "text/html":
                    charset = part.get_content_charset() or 'utf-8'
                    html_content += part.get_payload(decode=True).decode(charset, errors='ignore')
        else:
            if msg.get_content_type() == "text/plain":
                charset = msg.get_content_charset() or 'utf-8'
                encoding = msg.get('Content-Transfer-Encoding', '').lower()
                
                raw_payload = msg.get_payload(decode=True)
                if raw_payload:
                    if encoding == 'quoted-printable':
                        plain_text = quopri.decodestring(raw_payload).decode(charset, errors='ignore')
                    else:
                        plain_text = raw_payload.decode(charset, errors='ignore')
                    
                    # Clean up formatting
                    plain_text = plain_text.replace('=\n', '')
                    plain_text = re.sub(r'=([0-9A-F]{2})', lambda m: chr(int(m.group(1), 16)), plain_text)
                    
            elif msg.get_content_type() == "text/html":
                charset = msg.get_content_charset() or 'utf-8'
                html_content = msg.get_payload(decode=True).decode(charset, errors='ignore')
        
        return plain_text, html_content
    
    def parse_booking_info(self, text: str, msg: email.message.Message) -> BookingInfo:
        """Genius-level regex parser for Amadeus booking information."""
        booking = BookingInfo()
        
        # Extract passenger name from subject (decode if Base64 encoded)
        subject = str(msg.get('Subject', ''))
        if '=?UTF-8?B?' in subject:
            import base64
            # Decode Base64 encoded subject
            encoded_parts = re.findall(r'=\?UTF-8\?B\?([^?]+)\?=', subject)
            for encoded_part in encoded_parts:
                try:
                    decoded_part = base64.b64decode(encoded_part).decode('utf-8')
                    subject = subject.replace(f'=?UTF-8?B?{encoded_part}?=', decoded_part)
                except:
                    pass
        
        # Parse passenger name from subject: "LASTNAME/FIRSTNAME DATE ROUTE"
        name_match = re.search(r'([A-Z\s/]+?)\s+\d{1,2}[A-Z]{3}\d{4}', subject)
        if name_match:
            raw_name = name_match.group(1).strip()
            # Convert LASTNAME/FIRSTNAME to FIRSTNAME LASTNAME
            if '/' in raw_name:
                parts = raw_name.split('/')
                if len(parts) >= 2:
                    booking.passenger_name = f"{parts[1].strip()} {parts[0].strip()}"
                else:
                    booking.passenger_name = raw_name.replace('/', ' ')
            else:
                booking.passenger_name = raw_name
        
        # Extract booking reference - multiple patterns
        booking_ref_patterns = [
            r'BOOKING REF:\s*([A-Z0-9]{6,8})',
            r'FLIGHT BOOKING REF:\s*[A-Z]{2}/([A-Z0-9]{6,8})',
            r'REF:\s*([A-Z0-9]{6,8})'
        ]
        for pattern in booking_ref_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                booking.booking_ref = match.group(1)
                break
        
        # Extract booking date - multiple formats
        date_patterns = [
            r'DATE:\s*(\d{1,2}\s+[A-Z]+\s+\d{4})',
            r'DATE:\s*(\d{1,2}/\d{1,2}/\d{4})',
            r'(\d{1,2}\s+[A-Z]+\s+\d{4})'
        ]
        for pattern in date_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                booking.date = match.group(1).strip()
                break
        
        # Extract group/trip information (avoid matching "HAVE A GOOD TRIP")
        group_patterns = [
            r'GROUP\s+([^\r\n]+)',
            r'GROUP:\s*([^\r\n]+)',
            r'TRIP(?:\s+ID)?\s*:\s*([^\r\n]+)',
            r'GROUP\s+BOOKING:\s*([^\r\n]+)'
        ]
        for pattern in group_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                booking.group = match.group(1).strip()
                break
        
        # Extract ticket number with comprehensive pattern
        ticket_patterns = [
            r'TICKET:\s*([A-Z0-9/\s]+?)\s+FOR\s',
            r'ETKT\s+\d+\s+(\d+)\s+FOR',
            r'TICKET\s*[:\s]\s*([A-Z0-9/\s]+?)(?:\s+FOR|\r|\n)'
        ]
        for pattern in ticket_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                booking.ticket_number = match.group(1).strip()
                break
        
        return booking
    
    def parse_flights(self, text: str) -> List[FlightInfo]:
        """Genius-level regex parser for Amadeus flight information."""
        flights = []
        
        # Advanced multi-line flight section extraction including detail sections
        # This pattern captures flight header + its detail section (booking ref, baggage, meal, aircraft)
        flight_pattern = re.compile(
            r'FLIGHT\s+([A-Z0-9\s\-]+\s*-\s*[A-Z\s]+).*?(?=FLIGHT\s+[A-Z0-9\s\-]+\s*-\s*[A-Z\s]+|FLIGHT\(S\)\s+CALCULATED|GENERAL\s+INFORMATION|FLIGHT\s+TICKET|$)',
            re.DOTALL | re.IGNORECASE
        )
        
        flight_sections = flight_pattern.findall(text)
        full_matches = flight_pattern.finditer(text)
        
        for i, match in enumerate(full_matches):
            section = match.group(0)
            flight = FlightInfo()
            
            # Parse flight number and airline with advanced patterns
            flight_header_patterns = [
                r'FLIGHT\s+([A-Z]{2}\s*\d+)\s*-\s*([A-Z\s]+?)(?:\s+[A-Z]{3}\s+\d+\s+[A-Z]+\s+\d+)',
                r'FLIGHT\s+([A-Z]{2}\s*\d+)\s*-\s*([A-Z\s]+?)(?:\s+\w+\s+\d+)',
                r'FLIGHT\s+([A-Z]{2}\s*\d+)\s*-\s*([A-Z\s]+)'
            ]
            
            for pattern in flight_header_patterns:
                flight_match = re.search(pattern, section)
                if flight_match:
                    flight.flight_number = re.sub(r'\s+', ' ', flight_match.group(1).strip())
                    flight.airline = flight_match.group(2).strip()
                    break
            
            # Advanced departure parsing with multiple formats
            departure_patterns = [
                r'DEPARTURE:\s*([^,\(\r\n]+?)(?:,\s*([A-Z]{2}))?\s*\(([^)]+)\)\s*(\d{1,2}\s+[A-Z]{3}\s+\d{2}:\d{2})',
                r'DEPARTURE:\s*([^,\(\r\n]+?),?\s*([A-Z]{2})?\s*\(([^)]+)\)\s*(\d{1,2}\s+[A-Z]{3}\s+\d{2}:\d{2})',
                r'DEPARTURE:\s*([^\(\r\n]+?)\s*\(([^)]+)\)\s*(\d{1,2}\s+[A-Z]{3}\s+\d{2}:\d{2})',
                r'DEPARTURE:\s*([^\r\n]+?)\s*(\d{1,2}\s+[A-Z]{3}\s+\d{2}:\d{2})'
            ]
            
            for pattern in departure_patterns:
                dep_match = re.search(pattern, section)
                if dep_match:
                    groups = dep_match.groups()
                    flight.departure_city = groups[0].strip()
                    
                    if len(groups) >= 3:
                        # Patterns with airport info: groups[0]=city, groups[1]=country?, groups[2]=airport, groups[3]=datetime  
                        if len(groups) == 4 and groups[2]:  # Has airport info
                            flight.departure_airport = groups[2].strip()
                            datetime_str = groups[3]
                        else:
                            flight.departure_airport = groups[1].strip() if groups[1] else ""
                            datetime_str = groups[2] if len(groups) > 2 else groups[1]
                    else:
                        # Pattern 4: only city and datetime
                        datetime_str = groups[1]
                        
                    # Parse date and time
                    datetime_parts = datetime_str.strip().split(' ')
                    if len(datetime_parts) >= 3:
                        flight.departure_date = f"{datetime_parts[0]} {datetime_parts[1]}"
                        flight.departure_time = datetime_parts[2]
                    break
            
            # Advanced arrival parsing with multiple formats  
            arrival_patterns = [
                r'ARRIVAL:\s*([^,\(\r\n]+?)(?:,\s*([A-Z]{2}))?\s*\(([^)]+)\)(?:,\s*TERMINAL\s*\d+)?\s*(\d{1,2}\s+[A-Z]{3}\s+\d{2}:\d{2})',
                r'ARRIVAL:\s*([^,\(\r\n]+?),?\s*([A-Z]{2})?\s*\(([^)]+)\)(?:,\s*TERMINAL\s*\d+)?\s*(\d{1,2}\s+[A-Z]{3}\s+\d{2}:\d{2})',
                r'ARRIVAL:\s*([^\(\r\n]+?)\s*\(([^)]+)\)(?:,\s*TERMINAL\s*\d+)?\s*(\d{1,2}\s+[A-Z]{3}\s+\d{2}:\d{2})',
                r'ARRIVAL:\s*([^\r\n]+?)\s*(\d{1,2}\s+[A-Z]{3}\s+\d{2}:\d{2})'
            ]
            
            for pattern in arrival_patterns:
                arr_match = re.search(pattern, section)
                if arr_match:
                    groups = arr_match.groups()
                    flight.arrival_city = groups[0].strip()
                    
                    if len(groups) >= 3:
                        # Patterns with airport info: groups[0]=city, groups[1]=country?, groups[2]=airport, groups[3]=datetime  
                        if len(groups) == 4 and groups[2]:  # Has airport info
                            flight.arrival_airport = groups[2].strip()
                            datetime_str = groups[3]
                        else:
                            flight.arrival_airport = groups[1].strip() if groups[1] else ""
                            datetime_str = groups[2] if len(groups) > 2 else groups[1]
                    else:
                        # Pattern 4: only city and datetime (handles ASCII column formatting)
                        datetime_str = groups[1]
                        
                    # Parse date and time
                    datetime_parts = datetime_str.strip().split(' ')
                    if len(datetime_parts) >= 3:
                        flight.arrival_date = f"{datetime_parts[0]} {datetime_parts[1]}"
                        flight.arrival_time = datetime_parts[2]
                    break
            
            # Extract additional flight details with comprehensive patterns
            detail_patterns = {
                'booking_ref': [
                    r'FLIGHT BOOKING REF:\s*([A-Z0-9/]+)',
                    r'BOOKING REF:\s*([A-Z0-9/]+)'
                ],
                'class_type': [
                    r'RESERVATION CONFIRMED,\s*(ECONOMY)\s*\(([^)]+)\)',
                    r'(BUSINESS|ECONOMY|FIRST)\s*\(([^)]+)\)',
                    r'(ECONOMY|BUSINESS|FIRST)\s+CLASS'
                ],
                'duration': [
                    r'DURATION:\s*(\d{1,2}:\d{2})',
                    r'DURATION\s*(\d{1,2}H\s*\d{2}M?)',
                    r'FLIGHT TIME:\s*(\d{1,2}:\d{2})'
                ],
                'aircraft': [
                    r'EQUIPMENT:\s*([A-Z0-9\s\(\)-]+?)(?:\r?\n|$)',
                    r'AIRCRAFT:\s*([^\r\n]+?)(?:\r|\n|$)',
                    r'AC:\s*([^\r\n]+?)(?:\r|\n|$)'
                ],
                'meal': [
                    r'MEAL:\s*([A-Z\s/]+(?:\r?\n\s+[A-Z\s/]+)*?)(?=\r?\n\s*(?:NON\s+STOP|FLIGHT|$))',
                    r'MEAL:\s*([^\r\n]+?)(?:\r?\n|$)',
                    r'CATERING:\s*([^\r\n]+?)(?:\r|\n|$)'
                ],
                'baggage': [
                    r'BAGGAGE ALLOWANCE:\s*([A-Z0-9]+)',
                    r'BAGGAGE:\s*([A-Z0-9]+)',
                    r'BAG:\s*([A-Z0-9]+)'
                ]
            }
            
            # Apply all detail patterns
            for detail_type, patterns in detail_patterns.items():
                for pattern in patterns:
                    match = re.search(pattern, section, re.IGNORECASE)
                    if match:
                        if detail_type == 'booking_ref':
                            flight.booking_ref = match.group(1)
                        elif detail_type == 'class_type':
                            if len(match.groups()) >= 2 and match.group(2):
                                flight.class_type = f"{match.group(1).title()} ({match.group(2)})"
                            else:
                                flight.class_type = match.group(1).title()
                        elif detail_type == 'duration':
                            flight.duration = match.group(1)
                        elif detail_type == 'aircraft':
                            flight.aircraft = match.group(1).strip()
                        elif detail_type == 'meal':
                            flight.meal = match.group(1).strip()
                        elif detail_type == 'baggage':
                            flight.baggage = match.group(1).strip()
                        break
            
            # Only add flight if it has essential information
            if flight.flight_number or (flight.departure_city and flight.arrival_city):
                flights.append(flight)
        
        return flights
    
    def filter_valid_flights(self, flights: List[FlightInfo]) -> List[FlightInfo]:
        """Filter out flights with missing essential information."""
        valid_flights = []
        for flight in flights:
            # A flight is considered valid if it has flight number and at least departure/arrival info
            if (flight.flight_number and 
                flight.departure_city and 
                flight.arrival_city and
                flight.departure_date and 
                flight.arrival_date):
                valid_flights.append(flight)
        return valid_flights
    
    def format_company_header(self) -> str:
        """Format company header with logo and contact information."""
        return """
        <div class="company-header">
            <img src="https://www.triptojapan.com/logo.svg" alt="Trip to Japan" class="company-logo">
        </div>
        <div class="company-info">
            <div class="contact-grid">
                <div>
                    <div class="contact-item"><strong>Phone:</strong> +81 03-4578-2152</div>
                    <div class="contact-item"><strong>Email:</strong> info@triptojapan.com</div>
                </div>
                <div>
                    <div class="contact-item"><strong>Address:</strong><br>
                    Takanawa Travel K.K.,<br>
                    Kitashinagawa 5-11-1<br>
                    Shinagawa, Tokyo, Japan</div>
                </div>
            </div>
            <div style="margin-top: 8px; text-align: center; font-size: 9px; color: #666;">
                <strong>Certified Travel License</strong><br>
                Tokyo Metropolitan Government Office: No.3-8367
            </div>
        </div>
        """
    
    def format_booking_summary(self, booking: BookingInfo) -> str:
        """Format booking summary section."""
        return f"""
        <div class="booking-summary">
            <h2>Flight Booking Confirmation</h2>
            <div class="passenger-info">
                <div class="info-item">
                    <div class="info-label">Passenger</div>
                    <div>{html.escape(booking.passenger_name or 'N/A')}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Booking Reference</div>
                    <div>{html.escape(booking.booking_ref or 'N/A')}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Booking Date</div>
                    <div>{html.escape(booking.date or 'N/A')}</div>
                </div>
            </div>
            {f'<div class="info-item"><div class="info-label">Group</div><div>{html.escape(booking.group)}</div></div>' if booking.group else ''}
        </div>
        """
    
    def format_flight_card(self, flight: FlightInfo, flight_num: int) -> str:
        """Format a single flight as a card."""
        # Extract airport codes from city names
        dep_code = flight.departure_city.split(',')[0].strip() if flight.departure_city else 'N/A'
        arr_code = flight.arrival_city.split(',')[0].strip() if flight.arrival_city else 'N/A'
        
        # Try to extract 3-letter codes
        dep_code_match = re.search(r'\b([A-Z]{3})\b', dep_code)
        arr_code_match = re.search(r'\b([A-Z]{3})\b', arr_code)
        
        if dep_code_match:
            dep_airport_code = dep_code_match.group(1)
            dep_city_name = dep_code.replace(dep_airport_code, '').strip(' ()')
        else:
            dep_airport_code = dep_code[:3].upper() if len(dep_code) >= 3 else dep_code
            dep_city_name = flight.departure_airport or ''
            
        if arr_code_match:
            arr_airport_code = arr_code_match.group(1)
            arr_city_name = arr_code.replace(arr_airport_code, '').strip(' ()')
        else:
            arr_airport_code = arr_code[:3].upper() if len(arr_code) >= 3 else arr_code
            arr_city_name = flight.arrival_airport or ''
        
        return f"""
        <div class="flight-card">
            <div class="flight-header">
                Flight {flight.flight_number} - {flight.airline}
            </div>
            <div class="flight-route">
                <div class="airport-info">
                    <div class="airport-code">{dep_airport_code}</div>
                    <div class="airport-name">{html.escape(dep_city_name)}</div>
                    <div class="datetime">{flight.departure_date}</div>
                    <div class="datetime">{flight.departure_time}</div>
                </div>
                <div class="flight-arrow">→</div>
                <div class="airport-info">
                    <div class="airport-code">{arr_airport_code}</div>
                    <div class="airport-name">{html.escape(arr_city_name)}</div>
                    <div class="datetime">{flight.arrival_date}</div>
                    <div class="datetime">{flight.arrival_time}</div>
                </div>
            </div>
            <div class="flight-details">
                <div class="detail-item">
                    <div class="detail-label">Duration</div>
                    <div>{flight.duration or 'N/A'}</div>
                </div>
                <div class="detail-item">
                    <div class="detail-label">Class</div>
                    <div>{flight.class_type or 'N/A'}</div>
                </div>
                <div class="detail-item">
                    <div class="detail-label">Aircraft</div>
                    <div>{html.escape(flight.aircraft or 'N/A')}</div>
                </div>
                <div class="detail-item">
                    <div class="detail-label">Baggage</div>
                    <div>{flight.baggage or 'N/A'}</div>
                </div>
            </div>
            {f'<div style="padding: 10px 15px; font-size: 9px; border-top: 1px solid #000;"><strong>Meal:</strong> {html.escape(flight.meal)}</div>' if flight.meal else ''}
        </div>
        """
    
    def convert_to_html(self, msg: email.message.Message) -> str:
        """Convert email message to structured HTML format."""
        plain_text, html_content = self.extract_text_content(msg)
        
        # Use plain text for parsing as it's more reliable for structured data
        text_content = plain_text if plain_text else html_content
        
        if not text_content:
            text_content = "No readable content found in email."
        
        # Parse booking and flight information
        booking_info = self.parse_booking_info(text_content, msg)
        all_flights = self.parse_flights(text_content)
        flights = self.filter_valid_flights(all_flights)
        
        # Start building HTML
        html_doc = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>{self.css_style}</style>
        </head>
        <body>
        """
        
        # Add company logo only
        html_doc += """
        <div class="company-header">
            <img src="https://www.triptojapan.com/logo.svg" alt="Trip to Japan" class="company-logo">
        </div>
        """
        
        # Add booking summary
        html_doc += self.format_booking_summary(booking_info)
        
        # Add flight cards
        for i, flight in enumerate(flights, 1):
            html_doc += self.format_flight_card(flight, i)
        
        # Add ticket information if available
        if booking_info.ticket_number:
            html_doc += f"""
            <div class="ticket-info">
                <h3>Ticket Information</h3>
                <div><strong>Ticket Number:</strong> {html.escape(booking_info.ticket_number)}</div>
                <div><strong>Passenger:</strong> {html.escape(booking_info.passenger_name or 'N/A')}</div>
            </div>
            """
        
        # Add any additional notes or CO2 info if found
        co2_match = re.search(r'CO2 EMISSIONS IS ([\d.]+) KG/PERSON', text_content, re.IGNORECASE)
        if co2_match:
            html_doc += f"""
            <div style="border: 1px solid #000; padding: 10px; margin-top: 15px; font-size: 9px; text-align: center;">
                <strong>ENVIRONMENTAL IMPACT:</strong> Estimated CO2 emissions: {co2_match.group(1)} kg per person<br>
                Source: ICAO Carbon Emissions Calculator
            </div>
            """
        
        # Add company info at bottom
        html_doc += """
        <div class="company-info">
            <div class="contact-grid">
                <div>
                    <div class="contact-item"><strong>Phone:</strong> +81 03-4578-2152</div>
                    <div class="contact-item"><strong>Email:</strong> info@triptojapan.com</div>
                </div>
                <div>
                    <div class="contact-item"><strong>Address:</strong><br>
                    Takanawa Travel K.K.,<br>
                    Kitashinagawa 5-11-1<br>
                    Shinagawa, Tokyo, Japan</div>
                </div>
            </div>
            <div style="margin-top: 8px; font-size: 8px; color: #000;">
                <strong>Certified Travel License</strong><br>
                Tokyo Metropolitan Government Office: No.3-8367
            </div>
        </div>
        """
        
        html_doc += '</body></html>'
        
        return html_doc
    
    def convert_eml_to_pdf(self, eml_path: Path, output_path: Optional[Path] = None) -> Path:
        """Convert an EML file to PDF."""
        if not eml_path.exists():
            raise FileNotFoundError(f"EML file not found: {eml_path}")
        
        if output_path is None:
            output_path = eml_path.with_suffix('.pdf')
        
        console.print(f"Converting [bold blue]{eml_path.name}[/bold blue] to PDF...")
        
        # Parse the EML file
        msg = self.parse_eml_file(eml_path)
        
        # Convert to HTML
        html_content = self.convert_to_html(msg)
        
        # Generate PDF using WeasyPrint with safer settings
        try:
            import warnings
            warnings.filterwarnings('ignore')
            
            # Create HTML document with base_url to avoid network requests
            html_doc = weasyprint.HTML(string=html_content, base_url='.')
            
            # Write PDF with safer settings
            html_doc.write_pdf(output_path, presentational_hints=True)
            console.print(f"✓ Successfully created [bold green]{output_path.name}[/bold green]")
            return output_path
        except Exception as e:
            console.print(f"❌ Error converting {eml_path.name}: {e}")
            # Don't re-raise the exception to continue processing other files
            return None
    
    def batch_convert(self, input_dir: Path, output_dir: Optional[Path] = None) -> list[Path]:
        """Convert all EML files in a directory to PDF."""
        if not input_dir.exists() or not input_dir.is_dir():
            raise NotADirectoryError(f"Input directory not found: {input_dir}")
        
        if output_dir is None:
            output_dir = input_dir / "converted_pdfs"
        
        output_dir.mkdir(exist_ok=True)
        
        eml_files = list(input_dir.glob("*.eml"))
        if not eml_files:
            console.print(f"❌ No EML files found in {input_dir}")
            return []
        
        console.print(f"Found [bold]{len(eml_files)}[/bold] EML files to convert...")
        
        converted_files = []
        for eml_file in eml_files:
            try:
                output_file = output_dir / f"{eml_file.stem}.pdf"
                converted_file = self.convert_eml_to_pdf(eml_file, output_file)
                if converted_file:  # Only add if conversion was successful
                    converted_files.append(converted_file)
            except Exception as e:
                console.print(f"❌ Failed to convert {eml_file.name}: {e}")
                continue
        
        console.print(f"✓ Successfully converted [bold green]{len(converted_files)}[/bold green] files")
        return converted_files
    
    def recursive_batch_convert(self, input_dir: Path, output_dir: Optional[Path] = None) -> list[Path]:
        """Recursively convert all EML files in a directory tree to PDF, creating perfect mirror structure."""
        if not input_dir.exists() or not input_dir.is_dir():
            raise NotADirectoryError(f"Input directory not found: {input_dir}")
        
        # If no output directory specified, create one based on input directory name
        if output_dir is None:
            output_dir = Path("PDF") / input_dir.name
        
        # Ensure output directory exists
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Find all EML files recursively
        eml_files = list(input_dir.rglob("*.eml"))
        if not eml_files:
            console.print(f"❌ No EML files found recursively in {input_dir}")
            return []
        
        console.print(f"Found [bold]{len(eml_files)}[/bold] EML files recursively to convert...")
        console.print(f"Creating mirror structure: {input_dir} -> {output_dir}")
        
        converted_files = []
        
        # First, mirror the directory structure (including empty directories)
        for root, dirs, files in input_dir.walk():
            # Calculate relative path from input_dir
            relative_path = root.relative_to(input_dir)
            
            # Create corresponding directory in output
            output_subdir = output_dir / relative_path
            output_subdir.mkdir(parents=True, exist_ok=True)
        
        # Then convert all EML files while preserving structure
        for eml_file in eml_files:
            try:
                # Calculate relative path from input_dir to maintain structure
                relative_path = eml_file.relative_to(input_dir)
                
                # Create corresponding output path with PDF extension
                output_file = output_dir / relative_path.with_suffix('.pdf')
                
                converted_file = self.convert_eml_to_pdf(eml_file, output_file)
                if converted_file:  # Only add if conversion was successful
                    converted_files.append(converted_file)
                
            except Exception as e:
                console.print(f"❌ Failed to convert {eml_file.name}: {e}")
                continue
        
        console.print(f"✓ Successfully converted [bold green]{len(converted_files)}[/bold green] files recursively")
        console.print(f"✓ Mirror structure created at: [bold green]{output_dir}[/bold green]")
        return converted_files