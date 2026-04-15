#!/usr/bin/env python3
"""
Green Shield Report WO/UPRN Checker
Check all PDFs with job numbers starting with G-25 for Work Order or UPRN data
"""

import re
from pathlib import Path
import pdfplumber

print("\n" + "="*70)
print("GREEN SHIELD REPORT WO/UPRN CHECKER")
print("="*70 + "\n")

# Step 1: Find all PDFs with job numbers starting with G-25
print("Step 1: Searching for PDFs with job numbers starting with G-25...")
workspace_root = Path(r"c:\Users\Sherren\Desktop\feldman")
pdf_files = list(workspace_root.rglob("G-25*.pdf"))

if not pdf_files:
    print("  ⚠ No PDFs found with job numbers starting with G-25")
    print("  Searching in: " + str(workspace_root))
else:
    print(f"  ✓ Found {len(pdf_files)} PDF(s) with job numbers starting with G-25\n")

# Step 2: Extract WO/UPRN from each PDF
print("Step 2: Extracting WO/UPRN data from each report...\n")

results = []

for pdf_file in sorted(pdf_files):
    print(f"  Processing: {pdf_file.name}")
    
    try:
        with pdfplumber.open(pdf_file) as pdf:
            # Get first page to extract metadata
            first_page = pdf.pages[0]
            page_text = first_page.extract_text()
            
            # Extract Project Number (job number)
            project_no_match = re.search(r'Project No:\s*(G-\d+)', page_text, re.IGNORECASE)
            project_no = project_no_match.group(1) if project_no_match else "NOT FOUND"
            
            # Extract WO (Work Order)
            wo_match = re.search(r'WO:\s*(W\d+)', page_text, re.IGNORECASE)
            wo = wo_match.group(1) if wo_match else None
            
            # Extract UPRN (if no WO found)
            uprn_match = re.search(r'(?:UPRN|Project Ref):\s*(\d{4,5})', page_text, re.IGNORECASE)
            uprn = uprn_match.group(1) if uprn_match else None
            
            # Determine what was found
            has_identifier = wo or uprn
            identifier_type = "WO" if wo else ("UPRN" if uprn else "NONE")
            identifier_value = wo or uprn or "NOT FOUND"
            
            # Store result
            result = {
                'file': pdf_file.name,
                'project_no': project_no,
                'has_wo': bool(wo),
                'has_uprn': bool(uprn),
                'wo_value': wo,
                'uprn_value': uprn,
                'status': '✓ PASS' if has_identifier else '✗ FAIL'
            }
            results.append(result)
            
            # Print result
            status_icon = "✓" if has_identifier else "✗"
            print(f"    {status_icon} Project No: {project_no}")
            print(f"    {status_icon} Identifier: {identifier_type} = {identifier_value}")
            print()
            
    except Exception as e:
        print(f"    ✗ ERROR: {str(e)}\n")
        result = {
            'file': pdf_file.name,
            'project_no': "ERROR",
            'has_wo': False,
            'has_uprn': False,
            'wo_value': None,
            'uprn_value': None,
            'status': '✗ ERROR'
        }
        results.append(result)

# Step 3: Summary Report
print("="*70)
print("SUMMARY REPORT")
print("="*70 + "\n")

if results:
    print(f"{'File Name':<50} {'Project No':<12} {'WO/UPRN':<15} {'Status':<10}")
    print("-"*70)
    
    for r in results:
        identifier = r['wo_value'] or r['uprn_value'] or "N/A"
        identifier_type = ("WO:" if r['wo_value'] else "UPRN:") if (r['wo_value'] or r['uprn_value']) else ""
        print(f"{r['file']:<50} {r['project_no']:<12} {identifier_type} {identifier:<15} {r['status']:<10}")
    
    print("\n" + "-"*70)
    
    # Count results
    passed = len([r for r in results if '✓' in r['status']])
    failed = len([r for r in results if '✗' in r['status']])
    
    print(f"\nResults: {passed} PASS | {failed} FAIL")
    
    if failed == 0:
        print("✓ ALL PDFs have valid WO or UPRN identifiers!")
    else:
        print(f"⚠ {failed} PDF(s) missing WO/UPRN data - manual review needed")
else:
    print("No PDFs found to check")

print("\n" + "="*70 + "\n")
