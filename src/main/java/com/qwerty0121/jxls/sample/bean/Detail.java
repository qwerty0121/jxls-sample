package com.qwerty0121.jxls.sample.bean;

import java.math.BigDecimal;

import lombok.Data;

@Data
public class Detail {

  /** ID */
  private Integer id;

  /** 名前 */
  private String name;

  /** 金額 */
  private BigDecimal amount;

}
