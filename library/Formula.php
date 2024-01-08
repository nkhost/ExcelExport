<?php

namespace ExcelExport;

/**
 * Формула (в ячейке)
 */
class Formula
{
  protected $formula;
  
  public function __construct(string $formula)
  {
    $this->formula = $formula;
  }
  
  public function __toString()
  {
    return $this->formula;
  }
}
