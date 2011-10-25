require "poi_xls_rails/version"
require "csv"

module PoiXlsRails
  module SpreadsheetHelper
    
    def excel_document(opts={})
      download = opts.delete(:force_download)
      filename = opts.delete(:filename)
      template_path = opts.delete(:template_path)
      workbook = HSSFWorkbook.new
      yield(workbook)
      disposition(download, filename) if (download || filename)
      workbook
    end

    def disposition(download, filename)
      download = true if (filename && download == nil)
      disposition = download ? "attachment;" : "inline;"
      disposition += " filename=#{filename}" if filename
      headers["Content-Disposition"] = disposition
    end
  end

  class TemplateHandler
    # class_attribute :default_format
    # self.default_format = :xls
    
    def self.call(template)
      %Q{
        content_type = controller.content_type
        
        @file_extension = content_type == 'text/csv' ? 'csv' : 'xls'
        
        variables = controller.instance_variable_names
        variables -= %w[@template]
        
        if controller.respond_to?(:protected_instance_variables)
          variables -= controller.protected_instance_variables
        end
        
        variables.each do |name|
          instance_variable_set(name, controller.instance_variable_get(name))
        end
        
        workbook = #{template.source.strip}
        
        if content_type == 'text/csv'
          # http://en.wikipedia.org/wiki/Comma-separated_values#Specification
          
          workbook_arr = (0...workbook.number_of_sheets).inject([]) { |workbook_arr, sheet_index|
            sheet = workbook.getSheetAt(sheet_index)
            
            workbook_arr << (0..sheet.last_row_num).inject([]) { |sheet_arr, row_index|
              row = sheet.getRow(row_index)
              row_arr = []
              
              sheet_arr << (0..row.last_cell_num).inject([]) { |row_arr, cell_index|
                cell = row.getCell(cell_index)
                
                attempt_number = 0
                attempt_to_get_value = Proc.new do 
                  begin
                    case attempt_number
                    when 0
                      cell.getStringCellValue
                    when 1
                      cell.getNumericCellValue
                    when 2
                      cell.getRichStringCellValue
                    when 3
                      cell.getDateCellValue
                    when 4
                      cell.getErrorCellValue
                    end
                  rescue => e
                    raise e if attempt_number >= 5
                    attempt_number += 1
                    attempt_to_get_value.call
                  end
                end
                
                row_arr << attempt_to_get_value.call
                
              }.to_csv
            }.join("")
          }.join("\n\n")
        else
          outs = java.io.ByteArrayOutputStream.new
          workbook.write(outs)
          String.from_java_bytes(outs.toByteArray)
        end
      }
    end
  end
end

unless Mime::Type.lookup_by_extension :xls
  Mime::Type.register_alias "application/xls", :xls
end
unless Mime::Type.lookup_by_extension :csv
  Mime::Type.register_alias "text/csv", :csv
end
ActionView::Template.register_template_handler(:rxls, PoiXlsRails::TemplateHandler)
ActionView::Base.send(:include, PoiXlsRails::SpreadsheetHelper)
