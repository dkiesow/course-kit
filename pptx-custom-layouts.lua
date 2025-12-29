-- Lua filter for pandoc to apply custom PowerPoint slide layouts
-- This filter reads custom-style attributes and attempts to map them to slide layouts

-- Store layout mappings
local slide_layouts = {}
local current_slide_layout = nil

-- Function to read layout from div attributes
function Div(el)
  if el.attributes['custom-style'] then
    local layout_name = el.attributes['custom-style']
    -- Store this as the current slide's desired layout
    current_slide_layout = layout_name
    
    -- Mark this div as representing a slide layout preference
    el.attributes['pptx-layout'] = layout_name
  end
  return el
end

-- Function to handle headers and mark them with layout info
function Header(el)
  if current_slide_layout then
    -- Attach layout information to the header
    if not el.attributes then
      el.attributes = {}
    end
    el.attributes['slide-layout'] = current_slide_layout
  end
  return el
end

-- Reset layout tracking between slides
function HorizontalRule(el)
  current_slide_layout = nil
  return el
end

return {
  {Div = Div},
  {Header = Header},
  {HorizontalRule = HorizontalRule}
}
