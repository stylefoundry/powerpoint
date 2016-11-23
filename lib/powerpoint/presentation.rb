require 'zip/filesystem'
require 'fileutils'
require 'tmpdir'

module Powerpoint
  class Presentation
    include Powerpoint::Util

    attr_reader :slides, :masters, :layouts, :notes_masters, :themes, :rel_index, :layout_index, :theme_index, :extract_path, :master_rel

    def initialize
      @extract_path
      @dir
      @rel_index = 0
      @layout_index = 0
      @slides = []
      @masters = []
      @notes_masters = []
      @notes
      @layouts = []
      @themes = []
      @theme_index = 0
      @master_rel
      init_files
    end

    def add_intro(title, subtitile = nil)
      existing_intro_slide = @slides.select {|s| s.class == Powerpoint::Slide::Intro}[0]
      slide = Powerpoint::Slide::Intro.new(presentation: self, title: title, subtitile: subtitile)
      if existing_intro_slide
        @slides[@slides.index(existing_intro_slide)] = slide
      else
        @slides.insert 0, slide
      end
    end

    def add_textual_slide(title, content = [])
      @slides << Powerpoint::Slide::Textual.new(presentation: self, title: title, content: content)
    end

    def add_pictorial_slide(title, image_path, coords = {})
      @slides << Powerpoint::Slide::Pictorial.new(presentation: self, title: title, image_path: image_path, coords: coords)
    end

    def add_text_picture_slide(title, image_path, content = [])
      @slides << Powerpoint::Slide::TextPicSplit.new(presentation: self, title: title, image_path: image_path, content: content)
    end

    def add_picture_description_slide(title, image_path, content = [])
      @slides << Powerpoint::Slide::DescriptionPic.new(presentation: self, title: title, image_path: image_path, content: content)
    end

    def add_ff_trend_intro_slide(title, subtitle, image_path, coords = {})
      existing_intro_slide = @slides.select {|s| s.class == Powerpoint::Slide::Intro}[0]
      slide = Powerpoint::Slide::FFTrendIntro.new(presentation: self, title: title, subtitle: subtitle, image_path: image_path, coords: {})
      if existing_intro_slide
        @slides[@slides.index(existing_intro_slide)] = slide
      else
        @slides.insert 0, slide
      end
    end

    def add_ff_what_next_slide(title, content = {})
      @slides << Powerpoint::Slide::FFTrendWhatNext.new(presentation: self, title: title, content: content)
    end

    def add_ff_heading_text_slide(title, content)
      @slides << Powerpoint::Slide::FFTrendHeadingText.new(presentation: self, title: title, content: content)
    end

    def add_ff_three_row_text_slide(title, content)
      @slides << Powerpoint::Slide::FFTrendThreeRowText.new(presentation: self, title: title, content: content)
    end

    def add_ff_sector_impact_slide(title, content)
      @slides << Powerpoint::Slide::FFTrendSectorImpact.new(presentation: self, title: title, content: content)
    end

    def add_ff_associated_content_slide(title, subtitle, image_path, coords = {}, link_path = nil)
      @slides << Powerpoint::Slide::FFTrendIntro.new(presentation: self, title: title, subtitle: subtitle, image_path: image_path,  coords: {}, link_path: link_path)
    end

    def add_ff_embeded_slide(slide_content, slide_rel_content, images, charts, embeddings, notes, tags, drawings, master, notes_master, layout)
      @slides << Powerpoint::Slide::FFEmbededSlide.new(
        presentation: self, title: "", 
        content: slide_content, 
        rel_content: slide_rel_content, 
        images: images, 
        charts: charts, 
        embeddings: embeddings, 
        notes: notes, 
        tags: tags, 
        drawings: drawings, 
        master: master, 
        notes_master: notes_master,
        layout: layout
      )
    end

    def init_files
      @dir = Dir.mktmpdir
      @extract_path = "#{@dir}/extract_#{Time.now.strftime("%Y-%m-%d-%H%M%S")}"

      # Copy template to temp path
      FileUtils.copy_entry(TEMPLATE_PATH, extract_path)

      # Remove keep files
      Dir.glob("#{extract_path}/**/.keep").each do |keep_file|
        FileUtils.rm_rf(keep_file)
      end

      # Read in our template slide masters
      Dir.glob("#{extract_path}/ppt/slideMasters/*.xml").each do |master|
        @rel_index += 1
        master_rel_xml = Nokogiri::XML::Document.parse(File.open(master.gsub('ppt/slideMasters','ppt/slideMasters/_rels').gsub('.xml','.xml.rels')))
        theme_path = master_rel_xml.css('Relationship').select{ |node| node['Type'].include? 'relationships/theme'}.first['Target']
        @masters << { id: rel_index, file_path: master.gsub("#{extract_path}/ppt/slideMasters",'../slideMasters'), layouts: [], theme: theme_path }
        if !@themes.find{ |theme| theme[:file_path] == theme_path }
          @theme_index += 1
          @themes << { id: theme_index, file_path: theme_path }
        end
      end

      # Read in our notes template slide masters
      Dir.glob("#{extract_path}/ppt/notesMasters/*.xml").each do |master|
        @rel_index += 1
        @notes_masters << { id: rel_index, file_path: master.gsub("#{extract_path}/ppt/notesMasters",'../notesMasters') }
      end

      # Read in our layout templates
      Dir.glob("#{extract_path}/ppt/slideLayouts/*.xml").each do |layout|
        @layout_index += 1
        layout_rel_xml = Nokogiri::XML::Document.parse(File.open(layout.gsub('ppt/slideLayouts','ppt/slideLayouts/_rels').gsub('.xml','.xml.rels')))
        master_path = layout_rel_xml.css('Relationship').select{ |node| node['Type'].include? 'slideMaster'}.first['Target']
        new_layout = { id: layout_index, file_path: layout.gsub("#{extract_path}/ppt/slideLayouts",'../slideLayouts'), master: master_path }
        @layouts << new_layout
        @masters.find{ |master| master[:file_path] == master_path }[:layouts] << new_layout
      end
    end

    def save(path)
      # Save slides
      slides.each_with_index do |slide, index|
        slide.save(extract_path, index + 1)
      end

      # render master rels
      masters.each do |master_ref|
        @master_rel = master_ref
        render_view('slide_master.xml.rel.erb', extract_path + "/" + master_ref[:file_path].gsub('../slideMasters','ppt/slideMasters/_rels').gsub('.xml','.xml.rels'))
      end

      # render notes master rels
      notes_masters.each do |master|
        @master_rel = master
        render_view('notes_master.xml.rel.erb',extract_path + "/" + master[:file_path].gsub('../notesMasters','ppt/notesMasters/_rels').gsub('.xml','.xml.rels'))
      end

      # Render/save generic stuff
      render_view('content_type.xml.erb', "#{extract_path}/[Content_Types].xml")
      render_view('presentation.xml.rel.erb', "#{extract_path}/ppt/_rels/presentation.xml.rels")
      render_view('presentation.xml.erb', "#{extract_path}/ppt/presentation.xml")
      #render_view('view_props.xml.erb', "#{extract_path}/ppt/viewProps.xml")
      #render_view('table_styles.xml.erb', "#{extract_path}/ppt/tableStyles.xml")
      #render_view('pres_props.xml.erb', "#{extract_path}/ppt/presProps.xml")
      render_view('app.xml.erb', "#{extract_path}/docProps/app.xml")
      #render_view('core.xml.erb', "#{extract_path}/docProps/core.xml")

      # Create .pptx file
      File.delete(path) if File.exist?(path)
      Powerpoint.compress_pptx(extract_path, path)
      FileUtils.remove_entry @dir
      path
    end

    def file_types
      slides.map {|slide| slide.file_type if slide.respond_to? :file_type }.compact.flatten.uniq
    end

    def notes_slides
      slides.map {|slide| slide.notes_slides if slide.respond_to? :notes_slides }.compact.flatten
    end

    def add_master(xml, master_theme, master_layouts = [])
      @rel_index += 1
      File.open("#{extract_path}/ppt/slideMasters/slideMaster#{rel_index}.xml", "wb") do |f|
        f.write xml
      end
      new_master = { id: rel_index, file_path: "../slideMasters/slideMaster#{rel_index}.xml", layouts: master_layouts, theme: master_theme }
      @masters << new_master
      @masters.uniq!
      new_master
    end

    def add_notes_master(xml)
      @rel_index += 1
      File.open("#{extract_path}/ppt/notesMasters/notesMaster#{rel_index}.xml", "wb") do |f|
        f.write xml
      end
      new_master = { id: rel_index, file_path: "../notesMasters/notesMaster#{rel_index}.xml" }
      @notes_masters << new_master
      @notes_masters.uniq!
      new_master
    end

    def add_layout(xml, rel_xml, layout_master)
      @layout_index += 1
      File.open("#{extract_path}/ppt/slideLayouts/_rels/slideLayout#{layout_index}.xml.rels", "wb") do |f|
        f.write rel_xml
      end
      File.open("#{extract_path}/ppt/slideLayouts/slideLayout#{layout_index}.xml", "wb") do |f|
        f.write xml
      end
      new_layout = { id: layout_index, file_path: "../slideLayouts/slideLayout#{layout_index}.xml", master: layout_master }
      @masters.find{ |master_ref| master_ref[:file_path] == layout_master }[:layouts] << new_layout
      @layouts << new_layout
      @layouts.uniq!
      new_layout
    end

    def add_theme(xml)
      @theme_index += 1
      # write the theme xml with the correct file name
      File.open("#{extract_path}/ppt/theme/theme#{theme_index}.xml", "wb") do |f|
        f.write xml
      end
      new_theme = { id: theme_index, file_path: "../theme/theme#{theme_index}.xml" } 
      @themes << new_theme
      new_theme
    end

    def update_slide_masters
      masters.each do |master_ref|
        # open the xml
        master_xml = Nokogiri::XML::Document.parse(File.open("#{extract_path}/" + master_ref[:file_path].gsub('..','ppt')))
        master_xml.search('//p:sldLayoutIdLst').first.children.remove
        master_layouts = layouts.select{ |layout| layout[:master] == master_ref[:file_path] }
        master_layouts.each do |layout|
          node = Nokogiri::XML::Node.new 'p:sldLayoutId', master_xml
          node['r:id'] = "rId#{layout[:id]}"
          master_xml.search('//p:sldLayoutIdLst').first.add_child(node)
        end
        # save the file
        File.open("#{extract_path}/" + master_ref[:file_path].gsub('..','ppt'), 'wb') do |f|
          f.write master_xml.to_xml
        end
      end
    end
  end
end
