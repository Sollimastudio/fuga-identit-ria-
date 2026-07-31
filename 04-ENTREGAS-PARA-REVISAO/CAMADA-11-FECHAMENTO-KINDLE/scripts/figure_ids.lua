local figure_number = 0

function Figure(el)
  figure_number = figure_number + 1
  el.identifier = string.format("fig-%02d", figure_number)
  return el
end
