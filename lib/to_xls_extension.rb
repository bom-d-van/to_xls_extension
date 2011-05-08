# -*- encoding : utf-8 -*-
class Array
  def to_xls(options = {})
    output = 
      '<?xml version="1.0" encoding="UTF-8"?>
       <Workbook xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40" xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office">
       <Worksheet ss:Name="Sheet1"><Table>'

    if self.any?
      # create the only options if there are headers_mapping
      if options[:headers_mapping]
        options[:headers_mapping].to_options!
        options[:only] = options[:headers_mapping].keys
      end

      # --------- start create columns
      attributes = self.first.attributes.keys.sort.map(&:to_sym)
      if options[:only]
        columns = Array(options[:only]) & attributes
      else
        columns = attributes - Array(options[:except])
      end
      columns += Array(options[:methods])
      # --------- end create columns

      # create xls files
      if columns.any?
        unless options[:headers] == false
          output << "<Row>"
          if options[:headers_mapping]
            columns.each { |column| output << "<Cell><Data ss:Type=\"String\">#{options[:headers_mapping][column]}</Data></Cell>" }
          else
            klass = self.first.class
            columns.each { |column| output << "<Cell><Data ss:Type=\"String\">#{klass.human_attribute_name(column)}</Data></Cell>" }
          end
          output << "</Row>"
        end    

        self.each do |item|
          output << "<Row>"
          columns.each do |column|
            value = item.send column
            #output << "<Cell><Data ss:Type=\"#{value.is_a?(Integer) ? 'Number' : 'String'}\">#{value}</Data></Cell>"
            if value.is_a? Integer
              output << "<Cell><Data ss:Type=\"Number\">#{value}</Data></Cell>"
            elsif value.is_a? ActiveSupport::TimeWithZone
              output << "<Cell><Data ss:Type=\"String\">#{value.strftime("%Y-%m-%d")}</Data></Cell>"
            else
              output << "<Cell><Data ss:Type=\"String\">#{value}</Data></Cell>"
            end
          end
          output << "</Row>"
        end
      end
    end

    output << '</Table></Worksheet></Workbook>'
  end
end

