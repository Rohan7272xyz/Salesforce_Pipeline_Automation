def analyze_raw_data_sample(self, specific_file_path=None):
    """
    Analyze raw data file to understand column structure.
    
    Args:
        specific_file_path: If provided, analyze this specific file instead of most recent
    """
    try:
        if specific_file_path and Path(specific_file_path).exists():
            # Use the specific file provided
            latest_file = Path(specific_file_path)
            print(f"🔍 Analyzing SPECIFIC raw data file: {latest_file.name}")
        else:
            # Look for recent input files to analyze
            input_files = list(Config.INPUT_DIR.glob("pipeline_*.xlsx"))
            if not input_files:
                print("⚠️ No sample input files found for smart mapping analysis")
                return None
            
            # Use the most recent file
            latest_file = max(input_files, key=lambda f: f.stat().st_mtime)
            print(f"🔍 Analyzing raw data structure from: {latest_file.name}")
        
        workbook = load_workbook(latest_file, data_only=True)
        sheet = workbook.sheetnames[0]
        raw_sheet = workbook[sheet]
        
        # Get first few rows to find headers
        raw_values = []
        for row in raw_sheet.iter_rows(values_only=True, min_row=1, max_row=20):
            raw_values.append([str(cell).strip() if cell is not None else "" for cell in row])
        
        # Find header row (typically row 14, 0-indexed = 13)
        header_row_data = None
        for start_row in range(10, len(raw_values)):
            row = raw_values[start_row]
            row_text = " ".join(str(cell).lower() for cell in row if cell and cell.strip())
            
            # Look for key Salesforce column indicators
            indicators = ['capture', 'opportunity', 'salesforce', 'stage', 'positioning', 'govwin']
            matches = sum(1 for indicator in indicators if indicator in row_text)
            
            if matches >= 3:
                print(f"📋 Found raw data headers at row {start_row + 1}")
                header_row_data = row
                data_start_row = start_row + 1  # Data starts one row after headers
                break
        
        if not header_row_data:
            print("⚠️ Could not identify header row, using row 14 as fallback")
            header_row_data = raw_values[13] if len(raw_values) > 13 else None
            data_start_row = 14
        
        if not header_row_data:
            return None
        
        # Get sample data rows
        sample_data_rows = []
        for i in range(data_start_row, min(data_start_row + 5, len(raw_values))):
            if i < len(raw_values):
                sample_data_rows.append(raw_values[i])
        
        print(f"📊 Raw data structure analysis:")
        print(f"   Headers: {[h for h in header_row_data if h and h.strip()]}")
        print(f"   Sample data rows: {len(sample_data_rows)}")
        print(f"   Total columns found: {len([h for h in header_row_data if h and h.strip()])}")
        
        return {
            'headers': header_row_data,
            'sample_data': sample_data_rows,
            'data_start_row': data_start_row
        }
        
    except Exception as e:
        print(f"⚠️ Error analyzing raw data sample: {e}")
        import traceback
        print(traceback.format_exc())
        return None

def smart_column_mapping(self, template_columns, raw_data_info):
    """
    Create smart mapping between template columns and raw data columns.
    ENHANCED: Better validation and error handling for missing columns.
    """
    if not raw_data_info:
        print("❌ No raw data info available for mapping")
        return None
    
    raw_headers = raw_data_info['headers']
    sample_data = raw_data_info['sample_data']
    
    print(f"🔗 Smart Column Mapping Analysis:")
    print(f"   Template columns: {len(template_columns)}")
    print(f"   Raw data columns: {len(raw_headers)}")
    print(f"   Non-empty raw headers: {len([h for h in raw_headers if h and h.strip()])}")
    
    # Create mapping from template column to raw column index
    template_to_raw_mapping = {}
    unmapped_template_columns = []
    
    for template_col in template_columns:
        template_name = template_col['name']
        template_lower = template_name.lower().strip()
        
        best_match_index = None
        best_match_score = 0
        
        # Analyze each raw column
        for raw_idx, raw_header in enumerate(raw_headers):
            if not raw_header or not raw_header.strip():
                continue
                
            raw_header_lower = raw_header.lower().strip()
            
            # Score this column based on multiple factors
            score = 0
            
            # 1. Header name similarity
            if template_lower in raw_header_lower or raw_header_lower in template_lower:
                score += 10
            
            # 2. Specific keyword matching with higher scores for exact matches
            if 'positioning' in template_lower and 'positioning' in raw_header_lower:
                score += 25
            elif 'govwin' in template_lower and 'govwin' in raw_header_lower:
                score += 25
            elif 'capture' in template_lower and 'capture' in raw_header_lower:
                score += 20
            elif 'opportunity' in template_lower and 'opportunity' in raw_header_lower:
                score += 20
            elif ('sf' in template_lower or 'salesforce' in template_lower) and ('salesforce' in raw_header_lower or 'sf' in raw_header_lower):
                score += 25
            elif 'award' in template_lower and 'award' in raw_header_lower:
                score += 20
            elif 'ceiling' in template_lower and 'ceiling' in raw_header_lower:
                score += 20
            elif 'stage' in template_lower and 'stage' in raw_header_lower:
                score += 20
            
            # 3. Data content analysis (enhanced validation)
            if sample_data and score > 0:  # Only do expensive analysis if we have some match
                sample_values = [row[raw_idx] if raw_idx < len(row) else '' for row in sample_data[:3]]
                sample_values = [str(v).strip() for v in sample_values if v]
                
                if sample_values:
                    # Check data patterns to confirm mapping
                    if 'positioning' in template_lower:
                        # Should contain stage-like text (Sub, Capture, Qualification)
                        if any('sub' in v.lower() or 'capture' in v.lower() or 'qualification' in v.lower() for v in sample_values):
                            score += 30
                    elif 'govwin' in template_lower:
                        # Should contain numeric IDs
                        if any(v.isdigit() and len(v) >= 5 for v in sample_values):
                            score += 30
                    elif 'award' in template_lower or 'rfp' in template_lower:
                        # Should contain dates
                        if any('/' in v and any(c.isdigit() for c in v) for v in sample_values):
                            score += 25
                    elif 'ceiling' in template_lower or 'value' in template_lower:
                        # Should contain large numbers or currency
                        if any(v.replace(',', '').replace('$', '').replace('.', '').isdigit() for v in sample_values):
                            score += 20
            
            if score > best_match_score:
                best_match_score = score
                best_match_index = raw_idx
        
        # Only accept mappings with a reasonable confidence score
        if best_match_index is not None and best_match_score >= 10:
            template_to_raw_mapping[template_name] = {
                'raw_index': best_match_index,
                'raw_header': raw_headers[best_match_index],
                'score': best_match_score
            }
            print(f"   ✅ {template_name} → Raw Column {best_match_index + 1} '{raw_headers[best_match_index]}' (score: {best_match_score})")
        else:
            unmapped_template_columns.append(template_name)
            print(f"   ❌ No confident mapping found for: {template_name} (best score: {best_match_score})")
    
    # Report results
    print(f"\n📊 Mapping Results:")
    print(f"   Successfully mapped: {len(template_to_raw_mapping)} columns")
    print(f"   Failed to map: {len(unmapped_template_columns)} columns")
    
    if unmapped_template_columns:
        print(f"   Unmapped columns: {', '.join(unmapped_template_columns)}")
        print("   ⚠️ These columns will be skipped in processing")
    
    # Validate that we have enough mappings to proceed
    if len(template_to_raw_mapping) < len(template_columns) * 0.7:  # At least 70% mapped
        print(f"   ⚠️ WARNING: Only {len(template_to_raw_mapping)}/{len(template_columns)} columns mapped")
        print("   This may indicate a structural mismatch between template and raw data")
    
    return template_to_raw_mapping

# NEW: Add method to analyze with specific file
def analyze_and_update_with_file(self, input_file_path=None):
    """
    Main function to analyze template and generate new app.py.
    Enhanced to accept a specific input file for analysis.
    """
    try:
        print("🔍 Starting smart template analysis and app.py regeneration...")
        
        if input_file_path:
            print(f"📁 Using specific input file: {Path(input_file_path).name}")
        
        # Step 1: Extract headers from template
        headers = self.extract_template_headers()
        if not headers:
            raise ValueError("No headers found in template")
        
        # Step 2: Separate input columns from calendar columns
        input_columns, calendar_columns = self.identify_input_vs_calendar_columns(headers)
        
        if not input_columns:
            raise ValueError("No input data columns identified")
        
        # Step 3: Analyze raw data structure (with specific file if provided)
        raw_data_info = self.analyze_raw_data_sample(input_file_path)
        if not raw_data_info:
            print("⚠️ Warning: Could not analyze raw data, using template column order")
            column_mapping = None
        else:
            # Step 4: Create smart column mapping
            column_mapping = self.smart_column_mapping(input_columns, raw_data_info)
        
        # Step 5: Generate complete new app.py content
        new_app_content = self.generate_app_py_content(input_columns, calendar_columns, column_mapping)
        
        if not new_app_content:
            raise ValueError("Failed to generate app.py content")
        
        # Step 6: Backup current app.py
        self.backup_current_app_py()
        
        # Step 7: Write new app.py
        success = self.write_new_app_py(new_app_content)
        
        if success:
            print("🎉 Smart template analysis complete! New app.py generated.")
            print(f"📊 Processing Structure:")
            print(f"   Input data columns: {len(input_columns)} (from Salesforce)")
            print(f"   Calendar columns: {len(calendar_columns)} (template only)")
            print(f"   Total template columns: {len(input_columns + calendar_columns)}")
            
            if column_mapping:
                print("\n📋 Column mappings applied:")
                for template_col in input_columns:
                    if template_col['name'] in column_mapping:
                        mapping = column_mapping[template_col['name']]
                        print(f"   - {template_col['name']} ← Raw Column {mapping['raw_index'] + 1} '{mapping['raw_header']}'")
            
            return True
        else:
            print("❌ Failed to write new app.py")
            return False
            
    except Exception as e:
        print(f"❌ Template analysis failed: {e}")
        import traceback
        print(traceback.format_exc())
        return False

# Updated entry point function
def analyze_template_and_update_app(input_file_path=None):
    """
    Main entry point for template analysis - called by other scripts.
    Enhanced to accept a specific input file path.
    """
    analyzer = TemplateAnalyzer()
    if input_file_path:
        return analyzer.analyze_and_update_with_file(input_file_path)
    else:
        return analyzer.analyze_and_update()