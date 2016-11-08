require 'nokogiri'
require 'htmltoooxml'
require 'ruby_powerpoint'

include Htmltoooxml::XSLTHelper


# def html_to_ooxml(html)
#   source = Nokogiri::HTML(html.gsub(/>\s+</, '><'))
#   result = Htmltoooxml::Document.new().transform_doc_xml(source, false)
#   result.gsub!(/\s*<!--(.*?)-->\s*/m, '')
#   result = remove_declaration(result)
#   puts result
#   result
# end

def remove_whitespace(ooxml)
  ooxml.gsub(/\s+/, ' ').gsub(/>\s+</, '><').strip
end

def remove_declaration(ooxml)
  ooxml.sub(/<\?xml (.*?)>/, '').gsub(/\s*xmlns:(\w+)="(.*?)\s*"/, '')
end

