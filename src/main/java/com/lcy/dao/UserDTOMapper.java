package com.lcy.dao;

import java.util.List;

import com.lcy.model.UserDTO;

public interface UserDTOMapper {
    int deleteByPrimaryKey(Long id);

    int insert(UserDTO record);

    int insertSelective(UserDTO record);

    UserDTO selectByPrimaryKey(Long id);

    int updateByPrimaryKeySelective(UserDTO record);

    int updateByPrimaryKey(UserDTO record);
    
    List<UserDTO> selectAll();
}