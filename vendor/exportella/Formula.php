<?php

namespace Exportella;

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
