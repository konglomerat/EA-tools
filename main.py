#!/usr/bin/env python3

import eatools

# creds = eatools.get_credentials()
# drive = eatools.start_drive(creds)
# sheets = eatools.start_sheets(creds)

EA_ALL = eatools.EA_ALL()
EA_ALL.generate_pdf_of_bookings()
