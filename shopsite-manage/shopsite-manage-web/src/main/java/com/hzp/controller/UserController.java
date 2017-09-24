package com.hzp.controller;

import com.hzp.pojo.SysUser;

/**
 * Created by Administrator on 2017/9/24 0024.
 */
@Controller
@RequestMapping("/user")
public class UserController {
    @Resource
    private IUserService userService;

    @RequestMapping("/showUser")
    public String toIndex(HttpServletRequest request,Model model){
        int userId = Integer.parseInt(request.getParameter("id"));
        SysUser user = this.userService.getUserId(userId);
        model.addAttribute("user", user);
        return "showUser";
    }
}