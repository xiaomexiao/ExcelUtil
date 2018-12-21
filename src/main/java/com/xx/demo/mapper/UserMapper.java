package com.xx.demo.mapper;

import com.xx.demo.entity.User;
import com.baomidou.mybatisplus.mapper.BaseMapper;
import org.apache.ibatis.annotations.Mapper;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author admin
 * @since 2018-12-21
 */
@Mapper
public interface UserMapper extends BaseMapper<User> {

}
