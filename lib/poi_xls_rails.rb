require "poi_xls_rails/version"

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
    class_attribute :default_format
    self.default_format = :xls
    
    def self.call(template)
      %Q{
        variables = controller.instance_variable_names
        variables -= %w[@template]
        
        if controller.respond_to?(:protected_instance_variables)
          variables -= controller.protected_instance_variables
        end
        
        variables.each do |name|
          instance_variable_set(name, controller.instance_variable_get(name))
        end

        outs = java.io.ByteArrayOutputStream.new
        #{template.source.strip}.write(outs)
        String.from_java_bytes(outs.toByteArray)
      }
    end
  end
end

unless Mime::Type.lookup_by_extension :xls
  Mime::Type.register_alias "application/xls", :xls
end
ActionView::Template.register_template_handler(:rxls, PoiXlsRails::TemplateHandler)
ActionView::Base.send(:include, PoiXlsRails::SpreadsheetHelper)
