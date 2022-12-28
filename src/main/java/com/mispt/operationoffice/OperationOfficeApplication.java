package com.mispt.operationoffice;

import com.mispt.operationoffice.entity.OperateType;
import com.mispt.operationoffice.operate.IOperate;
import com.mispt.operationoffice.operate.OperateTemplate;
import java.util.Scanner;
import org.springframework.boot.WebApplicationType;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;

@SpringBootApplication
public class OperationOfficeApplication {

    public static void main(String[] args) {
//        SpringApplication.run(OperationOfficeApplication.class, args);
        SpringApplicationBuilder builder = new SpringApplicationBuilder(OperationOfficeApplication.class);
        builder.headless(false).web(WebApplicationType.SERVLET).run(args);

        // 1> 执行
        System.out.println("****************替换PPT/WORD内容******************");
        System.out.println("******* 请输入编码：0[PPT]，1[Word]，2[Excel] ");
        Scanner sca = new Scanner(System.in);
        int type = sca.nextInt();

        // 2> 解析
        OperateType operateType = OperateType.PPT;
        switch (type) {
            case 0:
                operateType = OperateType.PPT;
                break;
            case 1:
                operateType = OperateType.WORD;
                break;
            case 2:
                operateType = OperateType.EXCEL;
                break;
        }
        IOperate operate = OperateTemplate.getOperate(operateType);
        operate.exec(operateType);

        // 执行完成后退出程序
        System.exit(0);
    }

}
