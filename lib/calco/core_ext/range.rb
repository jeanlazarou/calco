class Range

  def as_grouping
    
    @grouping_range = true
    
    self
    
  end
  
  def grouping_range?
    ! @grouping_range.nil?
  end
  
end
