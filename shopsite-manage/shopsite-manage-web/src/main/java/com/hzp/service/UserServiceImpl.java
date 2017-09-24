package com.hzp.service;

import com.hzp.pojo.SysUser;

/**
 * Created by Administrator on 2017/9/24 0024.
 */

@Service("userService")
public class UserServiceImpl implements IUserService {
    @Resource
    private IUserDao userDao;
    @Override
    public SysUser getUserById(int userId) {
        // TODO Auto-generated method stub
        return this.userDao.selectByPrimaryKey(userId);
    }

}