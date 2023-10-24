require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'
require 'mimemagic'

module Powerpoint
  module Slide
    class FFEmbededSlide
      include Powerpoint::Util

      attr_reader :content
      attr_reader :rel_content
      attr_reader :title
      attr_reader :images
      attr_reader :charts
      attr_reader :embeddings
      attr_reader :theme_overrides
      attr_reader :notes
      attr_reader :tags
      attr_reader :drawings
      attr_reader :notes_slides
      attr_reader :master
      attr_reader :notes_master
      attr_reader :layouts
      attr_reader :layout
      attr_reader :file_types
      attr_reader :chart_images

      def initialize(options={})
        require_arguments [
          :presentation,
          :title, :content,
          :rel_content,
          :images,
          :charts,
          :embeddings,
          :notes,
          :tags,
          :drawings,
          :master,
          :notes_master,
          :layout,
          :theme_overrides,
          :chart_images
          ], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @file_types = []
        @notes_slides = []
        puts chart_images
      end

      def save(extract_path, index)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
        save_images(extract_path, index, images) if images && images.length> 0
        save_theme_overrides(extract_path, index) if theme_overrides && theme_overrides.length > 0
        save_charts(extract_path, index) if charts && charts.length > 0
        save_embeddings(extract_path, index) if embeddings && embeddings.length > 0
        save_images(extract_path,index, chart_images) if chart_images && chart_images.length > 0
        save_notes(extract_path, index) if notes && notes.length > 0
        save_tags(extract_path, index) if tags && tags.length > 0
        save_drawings(extract_path, index) if drawings && drawings.length > 0
      end

      def file_type
        file_types
      end

      private

      def save_rel_xml(extract_path, index)
        @index = index
        @tmp_content = rel_content.to_s
        @tmp_content.gsub!('charts',"charts/slide_#{@index}")
        @tmp_content.gsub!('media',"media/slide_#{@index}")
        @tmp_content.gsub!('../tags',"../tags/slide_#{@index}")
        xml = Nokogiri::XML::Document.parse @tmp_content
        xml.css('Relationship').select{ |node|
          if node['Type'].include? 'relationships/notesSlide'
            node['Target'] = "../notesSlides/notesSlide#{index}.xml"
          elsif node['Type'].include? 'relationships/masterSlide'
            node['Target'] = master[:file_path] if master && master[:file_path]
          elsif node['Type'].include? 'relationships/slideLayout'
            node['Target'] = layout[:file_path] if layout && layout[:file_path]
          end
        }
        @tmp_content = xml.to_s
        render_view('ff_embeded_slide_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end

      def save_slide_xml(extract_path, index)
        @index = index
        render_view('ff_embeded_slide_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end

      def save_images(extract_path, index, images)
        images.each do |image|
          FileUtils::mkdir_p "#{extract_path}/ppt/media/slide_#{index}"
          zip_entry = image.rewind
          file_path = zip_entry.name.to_s.gsub('media',"media/slide_#{index}").gsub('jpg','jpeg')
          begin
            File.open("#{extract_path}/" + file_path, 'wb') do |f|
              f.write zip_entry.get_input_stream.read
            end
            @file_types << { type: MimeMagic.by_magic(File.open("#{extract_path}/" + file_path)).type, path: "/#{file_path}" } unless file_path.include? "rels"
          rescue Exception => e
            puts "Error writing file #{e}"
          end
          image.close
        end
      end

      def save_charts(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/charts/slide_#{index}/_rels"
        charts.each do |chart|
          zip_entry = chart.rewind
          file_path = zip_entry.name.to_s.gsub('charts',"charts/slide_#{index}")
          if zip_entry.name.include? "rels"
            begin
              File.open("#{extract_path}/" +  file_path, 'wb') do |f|
                f.write zip_entry.get_input_stream.read
                  .gsub('../embeddings',"../../embeddings/slide_#{index}")
                  .gsub('../drawings',"../../drawings/slide_#{index}")
                  .gsub('../theme',"../../theme/slide_#{index}")
                  .gsub('../media',"../../media/slide_#{index}")
                  .gsub('smtClean="0"','')
              end
              rescue Exception => e
                puts "Error writing file #{e}"
              end
            else
              chart_xml = Nokogiri::XML::Document.parse zip_entry.get_input_stream.read
            #chart_xml.search('//c:chartSpace/c:lang').first.add_next_sibling('<c:style val="2"/>')
            #chart_xml.search('//c:chartSpace/c:externalData').first.add_child '<c:autoUpdate val="0"/>'

            data_label_xml = <<-EOXML
            <c:dLbls>
              <c:showLegendKey val="0"/>
              <c:showVal val="0"/>
              <c:showCatName val="0"/>
              <c:showSerName val="0"/>
              <c:showPercent val="0"/>
              <c:showBubbleSize val="0"/>
              <c:showLeaderLines val="0"/>
            </c:dLbls>
            EOXML
            bar_xml = chart_xml.search('//c:barChart/c:ser/c:spPr')
            if bar_xml.count > 0
              bar_xml.each do |node|
                #node.add_next_sibling(data_label_xml) unless chart_xml.search('//c:barChart/c:ser').count > 0
              end
            end

            cat_ax_xml = chart_xml.search('//c:catAx/c:delete')
            if cat_ax_xml.count < 1
              chart_xml.search('//c:catAx/c:scaling').each do |node|
                node.add_next_sibling '<c:delete val="0"/>'
              end
            end
            val_ax_xml = chart_xml.search('//c:valAx/c:delete')
            if val_ax_xml.count < 1
              chart_xml.search('//c:valAx/c:scaling').each do |node|
                node.add_next_sibling '<c:delete val="0"/>'
              end
            end
            begin
              File.open("#{extract_path}/" + file_path , 'wb:UTF-8') do |f|
                f.write chart_xml.to_xml.gsub('smtClean="0"','').strip
              end
            rescue Exception => e
              puts "Error writing file #{e}"
            end
            @file_types << { type: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" , path: "/#{file_path}" }
          end
          chart.close
        end
      end

      def save_embeddings(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/embeddings/slide_#{index}"
        embeddings.each do |embedding|
          begin
            zip_entry = embedding.rewind
            File.open("#{extract_path}/" + zip_entry.name.to_s.gsub('embeddings',"embeddings/slide_#{index}"), 'wb') do |f|
              f.write zip_entry.get_input_stream.read
            end
          embedding.close
          rescue Exception => e
            puts "Error writing file #{e}"
          end

        end
      end

      def save_drawings(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/drawings/slide_#{index}"
        drawings.each do |drawing|
          begin
            zip_entry = drawing.rewind
            file_path = zip_entry.name.to_s.gsub('drawings/',"drawings/slide_#{index}/")
            File.open("#{extract_path}/" + file_path, 'wb') do |f|
              f.write zip_entry.get_input_stream.read
                .gsub('smtClean="0"','')
            end
            @file_types << { type: "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml", path: "/#{file_path}" } unless file_path.include? "rels"
            drawing.close
          rescue Exception => e
            puts "Error writing file #{e}"
          end

        end
      end

      def save_theme_overrides(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/theme/slide_#{index}"
        theme_overrides.each do |theme_override|
          zip_entry = theme_override.rewind
          file_path = zip_entry.name.to_s.gsub('ppt/theme/',"ppt/theme/slide_#{index}/")
          begin
            File.open("#{extract_path}/" + file_path, 'wb') do |f|
              f.write zip_entry.get_input_stream.read
            end
          rescue Exception => e
            puts "Error writing file #{e}"
          end
          @file_types << { type: "application/vnd.openxmlformats-officedocument.themeOverride+xml", path: "/#{file_path}" } unless file_path.include? "rels"
          theme_override.close
        end
      end

      def save_tags(extract_path, index)
        FileUtils::mkdir_p "#{extract_path}/ppt/tags/slide_#{index}"
        tags.each do |tag|
          zip_entry = tag.rewind
          file_path = zip_entry.name.to_s.gsub('ppt/tags/',"ppt/tags/slide_#{index}/")
          begin
            File.open("#{extract_path}/" + file_path, 'wb') do |f|
              f.write zip_entry.get_input_stream.read
            end
          rescue Exception => e
            puts "Error writing file #{e}"
          end
          @file_types << { type: "application/vnd.openxmlformats-officedocument.presentationml.tags+xml", path: "/#{file_path}" } unless file_path.include? "rels"
          tag.close
        end
      end

      def save_notes(extract_path, index)
       FileUtils::mkdir_p "#{extract_path}/ppt/notesSlides/_rels"
        notes.each do |note|
          zip_entry = note.rewind
          if zip_entry.name.include? "rels"
            notes_xml = Nokogiri::XML::Document.parse zip_entry.get_input_stream.read
            notes_xml.css('Relationship').select{ |node|
              if node['Type'].include? 'relationships/slide'
                node['Target'] = "../slides/slide#{index}.xml"
              elsif node['Type'].include? 'relationships/notesMaster'
                node['Target'] = notes_master[:file_path]
              end
            }
            begin
              File.open("#{extract_path}/ppt/notesSlides/_rels/notesSlide#{index}.xml.rels", 'wb') do |f|
                f.write notes_xml
              end
            rescue Exception => e
              puts "Error writing file #{e}"
            end
          else
            file_path = "ppt/notesSlides/notesSlide#{index}.xml"
            @notes_slides << { type: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml", path: "#{file_path}" }
            notes_xml =  Nokogiri::XML::Document.parse  zip_entry.get_input_stream.read
            begin
              File.open("#{extract_path}/" + file_path , 'wb') do |f|
                f.write notes_xml.to_xml.gsub('smtClean="0"','')
              end
            rescue Exception => e
              puts "Error writing file #{e}"
            end
          end
          note.close
        end
      end
    end
  end
end
